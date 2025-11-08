const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

// Fonction pour demander à l'utilisateur de choisir un fichier
async function promptForFile() {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    return new Promise((resolve) => {
        rl.question('Entrez le chemin du fichier Excel (ou appuyez sur Entrée pour utiliser le fichier par défaut): ', (answer) => {
            rl.close();
            resolve(answer.trim() || 'report_standard_S112Part1.xlsx');
        });
    });
}

// Fonction pour lire le fichier Excel et extraire les données
function extractDataFromExcel(filePath) {
    try {
        // Vérifier si le fichier existe
        if (!fs.existsSync(filePath)) {
            throw new Error(`Le fichier ${filePath} n'existe pas`);
        }

        // Lire le fichier Excel
        const workbook = XLSX.readFile(filePath);

        // Obtenir la première feuille
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Convertir la feuille en JSON pour faciliter la manipulation
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        console.log('\n=== Aperçu du fichier Excel ===');
        console.log('Feuille:', sheetName);
        console.log('Nombre de lignes de données:', jsonData.length - 1);

        // La première ligne contient les en-têtes
        const headers = jsonData[0];

        // Trouver les indices des colonnes pour F5, F15, F25, F35
        const columnIndices = {};
        const targets = ['F5', 'F15', 'F25', 'F35'];

        targets.forEach(target => {
            columnIndices[target] = {
                equidose: -1,
                uncertainty: -1
            };

            headers.forEach((header, index) => {
                if (header && typeof header === 'string') {
                    // Chercher "F5 Gamma EquiDose", "F5 Gamma Uncertainty", etc.
                    if (header.includes(target) && header.includes('EquiDose')) {
                        columnIndices[target].equidose = index;
                    }
                    if (header.includes(target) && header.includes('Uncertainty')) {
                        columnIndices[target].uncertainty = index;
                    }
                }
            });
        });

        console.log('\n=== Colonnes trouvées ===');
        targets.forEach(target => {
            console.log(`${target}:`);
            console.log(`  EquiDose: colonne ${columnIndices[target].equidose} - ${headers[columnIndices[target].equidose]}`);
            console.log(`  Uncertainty: colonne ${columnIndices[target].uncertainty} - ${headers[columnIndices[target].uncertainty]}`);
        });

        // Extraire les données pour chaque ligne (à partir de la ligne 2)
        const results = [];

        for (let rowIndex = 1; rowIndex < jsonData.length; rowIndex++) {
            const row = jsonData[rowIndex];
            const rowData = {
                rowNumber: rowIndex + 1,
                irn: row[0] || 'N/A',
                data: []
            };

            targets.forEach(target => {
                const equidoseIndex = columnIndices[target].equidose;
                const uncertaintyIndex = columnIndices[target].uncertainty;

                rowData.data.push({
                    target: target,
                    equidose: equidoseIndex >= 0 ? (row[equidoseIndex] || 'N/A') : 'N/A',
                    uncertainty: uncertaintyIndex >= 0 ? (row[uncertaintyIndex] || 'N/A') : 'N/A'
                });
            });

            results.push(rowData);
        }

        console.log('\n=== Extraction des valeurs ===');
        results.forEach(rowData => {
            console.log(`\nLigne ${rowData.rowNumber} (IRN: ${rowData.irn}):`);
            rowData.data.forEach(item => {
                console.log(`  ${item.target}: EquiDose = ${item.equidose}, Uncertainty = ${item.uncertainty}`);
            });
        });

        return results;

    } catch (error) {
        console.error('Erreur lors de la lecture du fichier Excel:', error.message);
        throw error;
    }
}

// Fonction pour générer le fichier texte avec les résultats
function generateTextFile(results, outputPath) {
    let content = '=== RÉSULTATS D\'EXTRACTION ===\n\n';

    results.forEach((rowData, index) => {
        content += `────────────────────────────────────────\n`;
        content += `Ligne ${rowData.rowNumber} - IRN: ${rowData.irn}\n`;
        content += `────────────────────────────────────────\n\n`;

        rowData.data.forEach(item => {
            content += `Valeur de ${item.target}:\n`;
            content += `  EquiDose:     ${item.equidose}\n`;
            content += `  Incertitude:  ${item.uncertainty}\n`;
            content += `\n`;
        });

        content += '\n';
    });

    // Ajouter un tableau récapitulatif
    content += '\n═══════════════════════════════════════════════════════════════\n';
    content += 'TABLEAU RÉCAPITULATIF\n';
    content += '═══════════════════════════════════════════════════════════════\n\n';

    // En-tête du tableau
    content += 'IRN'.padEnd(15) + ' | ';
    content += 'Valeur'.padEnd(8) + ' | ';
    content += 'F5'.padEnd(15) + ' | ';
    content += 'F15'.padEnd(15) + ' | ';
    content += 'F25'.padEnd(15) + ' | ';
    content += 'F35'.padEnd(15) + '\n';
    content += '─'.repeat(15) + '─┼─';
    content += '─'.repeat(8) + '─┼─';
    content += '─'.repeat(15) + '─┼─';
    content += '─'.repeat(15) + '─┼─';
    content += '─'.repeat(15) + '─┼─';
    content += '─'.repeat(15) + '\n';

    results.forEach(rowData => {
        // Ligne EquiDose
        content += String(rowData.irn).padEnd(15) + ' | ';
        content += 'EquiDose'.padEnd(8) + ' | ';
        rowData.data.forEach(item => {
            content += String(item.equidose).padEnd(15) + ' | ';
        });
        content += '\n';

        // Ligne Incertitude
        content += ''.padEnd(15) + ' | ';
        content += 'Incert.'.padEnd(8) + ' | ';
        rowData.data.forEach(item => {
            content += String(item.uncertainty).padEnd(15) + ' | ';
        });
        content += '\n';

        content += ''.padEnd(15) + ' | ';
        content += ''.padEnd(8) + ' | ';
        content += ''.padEnd(15) + ' | ';
        content += ''.padEnd(15) + ' | ';
        content += ''.padEnd(15) + ' | ';
        content += ''.padEnd(15) + '\n';
    });

    fs.writeFileSync(outputPath, content, 'utf8');
    console.log(`\n✓ Fichier texte généré: ${outputPath}`);
}

// Fonction principale
async function main() {
    console.log('=== Extracteur de données Excel ===\n');

    // Demander le fichier à l'utilisateur
    const filePath = await promptForFile();

    console.log(`\nFichier sélectionné: ${filePath}`);

    // Extraire les données
    const results = extractDataFromExcel(filePath);

    // Générer le fichier de sortie
    const outputPath = 'resultats_extraction.txt';
    generateTextFile(results, outputPath);

    console.log('\n✓ Extraction terminée avec succès!');
}

// Exécuter le script
main().catch(error => {
    console.error('Erreur:', error.message);
    process.exit(1);
});
