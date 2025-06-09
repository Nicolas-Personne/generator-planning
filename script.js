// Sélection des éléments HTML
const button = document.querySelector("button");
const file = document.querySelector("input[type='file']");
const classesContainer = document.querySelector("#classes-container");
const dateInput = document.querySelector("input[type='date']");
const planningContainer = document.querySelector("#planning");
const headerPlanningContainer = document.querySelector("#header-container");
const tBodyContainer = document.querySelector("tbody");

// Variables de travail
const classes = []; // stocke les classes récupérées depuis la feuille "Data"
const includedClasses = []; // stocke les classes sélectionnées par l'utilisateur
let intervenantsArray = {}; // dictionnaire des enseignants
let subjectArray = {}; // dictionnaire des matières

// Tableaux utilitaires
const months = [
	"jan",
	"fev",
	"mar",
	"avr",
	"mai",
	"juin",
	"juil",
	"aout",
	"sept",
	"oct",
	"nov",
	"dec",
];
const days = ["LUN", "MAR", "MER", "JEU", "VEN"];

// Lorsqu’un fichier est chargé
file.addEventListener("change", (e) => {
	const uploadedFile = file.files[0];
	const reader = new FileReader();

	reader.onload = function (e) {
		const data = new Uint8Array(e.target.result);
		const workbook = XLSX.read(data, { type: "array" });

		workbook.SheetNames.map((sheet) => {
			if (sheet === "Data") {
				const worksheet = workbook.Sheets[sheet];
				const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 0 });

				// Récupère la couleur (classe) pour les checkbox
				jsonData.map((item) => {
					if (item.Couleurs) classes.push(item.Couleurs);
				});
			}
		});

		// Crée dynamiquement les checkbox pour chaque classe
		classes.map((item) => {
			const div = document.createElement("div");
			const input = document.createElement("input");
			input.type = "checkbox";
			input.name = "classes[]";
			input.value = item;
			input.id = item;

			const label = document.createElement("label");
			label.innerText = item;
			label.htmlFor = item;

			div.appendChild(input);
			div.appendChild(label);
			classesContainer.appendChild(div);
		});
	};

	reader.readAsArrayBuffer(uploadedFile);
});

// Convertisseur de date Excel → JS
const excelDateToJSDate = (serial) => {
	const utc_days = Math.floor(serial - 25569); // Excel date serial à UTC
	const utc_value = utc_days * 86400;
	return new Date(utc_value * 1000);
};

// Convertisseur de date JS → Excel
function dateToExcelSerial(date) {
	const msPerDay = 86400000;
	const epoch1900 = Date.UTC(1899, 11, 30);
	return (date.getTime() - epoch1900) / msPerDay;
}

// Remplit le planning HTML et le convertit en PDF
async function fillPlanning(fullDataArray, startDate, endDate, classe) {
	headerPlanningContainer.innerHTML = `
    <div id="header">
      <img src="./logo.png" alt="" />
      <h1>${classe}</h1>
    </div>
    <h2>
      Planning du <span id="start-date">${startDate}</span>
      au <span id="end-date">${endDate}</span>
    </h2>`;

	let k = 0;
	for (key in fullDataArray) {
		tBodyContainer.innerHTML += `
      <tr>
        <td rowspan="3" id="day">${days[k]}</td>
        <td class="head-row">${key}</td>
        <td class="head-row">${
					classe.startsWith("B") && ["1", "2"].includes(classe[1])
						? "9:00 - 12:00"
						: "9:00 - 12:30"
				}</td>
        <td class="head-row">${
					classe.startsWith("B") && ["1", "2"].includes(classe[1])
						? "13:00 - 16:00"
						: "13:30 - 17:00"
				}</td>
      </tr>
      <tr>
        <td class="side">Enseignant</td>
        <td class="content">${fullDataArray[key].enseignantAm ?? ""}</td>
        <td class="content">${fullDataArray[key].enseignantPm ?? ""}</td>
      </tr>
      <tr>
        <td class="side">Matière</td>
        <td class="content">${fullDataArray[key].matiereAm ?? ""}</td>
        <td class="content">${fullDataArray[key].matierePm ?? ""}</td>
      </tr>
    `;
		k++;
	}

	// Attente pour le rendu HTML avant la génération PDF
	await new Promise((resolve) => setTimeout(resolve, 300));

	// Génère et télécharge le PDF
	await html2pdf()
		.set({
			margin: 0.5,
			filename: `[${classe}]Planning du ${startDate} au ${endDate}.pdf`,
			image: { type: "jpeg", quality: 0.98 },
			html2canvas: { scale: 2 },
			jsPDF: { unit: "in", format: "a4", orientation: "landscape" },
		})
		.from(planningContainer)
		.save();

	// Nettoyage du DOM
	headerPlanningContainer.innerHTML = "";
	tBodyContainer.innerHTML = "";
}

// Réorganise une feuille Excel transposée pour les intervenants
function transposeSheet(data) {
	const result = {};
	const headers = data[0];

	headers.forEach((header, colIndex) => {
		if (header === "Intervenants Interne") {
			result[`${header}-lastname`] = data.slice(1).map((row) => row[colIndex]);
			result[`${header}-firstname`] = data
				.slice(1)
				.map((row) => row[colIndex + 1]);
			result[`${header}-initial`] = data
				.slice(1)
				.map((row) => row[colIndex - 1]);
		} else {
			result[header] = data.slice(1).map((row) => row[colIndex]);
		}
	});

	return result;
}

// Charge les matières par classe
const loadSubject = (json) => {
	const result = {};
	json.forEach((item, index) => {
		if (index > 1 && item[1] && item[2]) {
			result[item[1]] = item[2];
		}
	});
	return result;
};

// Charge les intervenants internes
const loadInter = (json) => {
	let sheetInter = transposeSheet(json);

	sheetInter["Intervenants Interne-initial"].map((item, index) => {
		if (item) {
			intervenantsArray = {
				...intervenantsArray,
				[item]:
					(sheetInter["Intervenants Interne-firstname"][index] ?? "Autonomie") +
					" " +
					(sheetInter["Intervenants Interne-lastname"][index]?.toUpperCase() ??
						""),
			};
		}
	});
};

// Génère les PDF après avoir sélectionné les classes et la date
async function generatePDF(e) {
	e.preventDefault();
	const classesInput = document.querySelectorAll("input[name='classes[]'");

	let j = 2;
	for (let i = 0; i < classesInput.length; i++) {
		if (classesInput[i].checked) {
			includedClasses.push({ classe: classesInput[i].value, interval: j });
		}
		j += 5;
	}

	const uploadedFile = file.files[0];
	const reader = new FileReader();

	reader.onload = async function (e) {
		const data = new Uint8Array(e.target.result);
		const workbook = XLSX.read(data, { type: "array" });

		// Chargement des matières
		workbook.SheetNames.map((sheet) => {
			includedClasses.map((item) => {
				const sheetName = `P_${item.classe.split(" ")[0]}${
					item.classe.split(" ")[1]
						? "_" + item.classe.split(" ")[1].toUpperCase()
						: ""
				}`;
				if (sheet === sheetName) {
					const worksheet = workbook.Sheets[sheet];
					const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
					const result = loadSubject(jsonData);
					subjectArray = { ...subjectArray, [item.classe]: result };
				}
			});
		});

		// Chargement des intervenants et du planning
		workbook.SheetNames.map((sheet) => {
			if (sheet === "Data") {
				const worksheet = workbook.Sheets[sheet];
				const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
				loadInter(jsonData);
			}

			if (sheet === "Planning_2025_2026") {
				const worksheet = workbook.Sheets[sheet];
				const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

				const choosenDate = new Date(dateInput.value);
				const choosenSerial = dateToExcelSerial(choosenDate);

				let fullDataArray = {};
				let startDate = "";
				let endDate = "";

				jsonData.map(async (item, index) => {
					let indexOfDate = item.indexOf(choosenSerial);
					if (indexOfDate !== -1) {
						let indexOfRow = index;

						includedClasses.map((classe) => {
							for (let i = 0; i < 5; i++) {
								let newDate = new Date(choosenDate.getTime() + i * 86400000);
								let changingMonth =
									newDate.getMonth() !== choosenDate.getMonth();
								if (changingMonth) {
									indexOfRow += 63;
									indexOfDate = 1;
								}

								// Formatage des dates
								if (i === 0) {
									startDate = `${String(newDate.getDate()).padStart(
										2,
										"0"
									)}-${String(newDate.getMonth() + 1).padStart(
										2,
										"0"
									)}-${newDate.getFullYear()}`;
								}
								if (i === 4) {
									endDate = `${String(newDate.getDate()).padStart(
										2,
										"0"
									)}-${String(newDate.getMonth() + 1).padStart(
										2,
										"0"
									)}-${newDate.getFullYear()}`;
								}

								fullDataArray[classe.classe] = {
									...fullDataArray[classe.classe],
									[`${newDate.getDate()}-${months[newDate.getMonth()]}`]: {
										enseignantAm:
											intervenantsArray[
												jsonData[indexOfRow + classe.interval]?.[
													indexOfDate + i
												]
											],
										enseignantPm:
											intervenantsArray[
												jsonData[indexOfRow + classe.interval + 3]?.[
													indexOfDate + i
												]
											],
										matiereAm:
											subjectArray[classe.classe]?.[
												jsonData[indexOfRow + classe.interval + 1]?.[
													indexOfDate + i
												]
											],
										matierePm:
											subjectArray[classe.classe]?.[
												jsonData[indexOfRow + classe.interval + 2]?.[
													indexOfDate + i
												]
											],
									},
								};
							}
						});

						// Génère un PDF par classe
						for (const classe in fullDataArray) {
							await fillPlanning(
								fullDataArray[classe],
								startDate,
								endDate,
								classe
							);
						}
					}
				});
			}
		});
	};
	reader.readAsArrayBuffer(uploadedFile);
}

// Clique sur le bouton = génération du planning PDF
button.addEventListener("click", (e) => generatePDF(e));
