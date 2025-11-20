// URL du webhook Power Automate pour LIRE les données
const WEBHOOK_READ_URL = "https://defaultefdc051fe1144edaa7c5efec56da13.b1.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/68fb07af56e94845b714ce22d00b5f4c/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=hTGYXli6b3lv0VdHzQD8lKdAc07jyjerdzHloXKOV4c";

// Variables globales
let donneesCompletes = [];
let donneesFiltrees = [];
let charts = {};

// Charger les données depuis Power Automate
async function chargerDonnees() {
    const loadingMessage = document.getElementById('loadingMessage');
    const errorMessage = document.getElementById('errorMessage');
    
    loadingMessage.classList.add('show');
    errorMessage.style.display = 'none';
    
    try {
        const response = await fetch(WEBHOOK_READ_URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            }
        });
        
        if (!response.ok) {
            throw new Error(`Erreur HTTP: ${response.status}`);
        }
        
        const data = await response.json();
        
        // Les données sont dans la propriété "value" de la réponse Excel
        donneesCompletes = data.value || data || [];
        donneesFiltrees = [...donneesCompletes];
        
        console.log('✅ Données chargées:', donneesCompletes.length, 'lignes');
        
        // Afficher les données
        afficherDashboard();
        
    } catch (error) {
        console.error('❌ Erreur lors du chargement:', error);
        errorMessage.style.display = 'block';
    } finally {
        loadingMessage.classList.remove('show');
    }
}

// Afficher le dashboard complet
function afficherDashboard() {
    if (donneesFiltrees.length === 0) {
        document.getElementById('errorMessage').style.display = 'block';
        document.getElementById('errorMessage').innerHTML = '<p>ℹ️ Aucune donnée disponible. Commencez par saisir des données.</p>';
        return;
    }
    
    calculerKPIs();
    afficherGraphiques();
    afficherTableau();
}

// Calculer les KPIs
function calculerKPIs() {
    let totalProdBonne = 0;
    let totalRebuts = 0;
    let totalTempsArret = 0;
    
    donneesFiltrees.forEach(ligne => {
        totalProdBonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
        totalRebuts += parseInt(ligne['Rebuts'] || 0);
        
        // Calculer le temps d'arrêt total
        const equipDuree = parseInt(ligne['Équipement Durée'] || ligne['Équipement_x0020_Durée'] || 0);
        const qualiteDuree = parseInt(ligne['Qualité Durée'] || ligne['Qualité_x0020_Durée'] || 0);
        const orgDuree = parseInt(ligne['Organisation Durée'] || ligne['Organisation_x0020_Durée'] || 0);
        const autresDuree = parseInt(ligne['Autres Durée'] || ligne['Autres_x0020_Durée'] || 0);
        
        totalTempsArret += equipDuree + qualiteDuree + orgDuree + autresDuree;
    });
    
    const totalProduction = totalProdBonne + totalRebuts;
    const tauxRebuts = totalProduction > 0 ? ((totalRebuts / totalProduction) * 100).toFixed(2) : 0;
    
    // Afficher les KPIs
    document.getElementById('kpiProdTotal').textContent = totalProduction.toLocaleString();
    document.getElementById('kpiProdBonne').textContent = totalProdBonne.toLocaleString();
    document.getElementById('kpiRebuts').textContent = totalRebuts.toLocaleString();
    document.getElementById('kpiTauxRebuts').textContent = tauxRebuts + '%';
    document.getElementById('kpiTempsArret').textContent = totalTempsArret + ' min';
}

// Afficher les graphiques
function afficherGraphiques() {
    afficherGraphiqueProductionHeure();
    afficherGraphiqueArrets();
    afficherGraphiqueProductionEquipe();
    afficherGraphiqueDureeArrets();
}

// Graphique: Production par heure
function afficherGraphiqueProductionHeure() {
    const ctx = document.getElementById('chartProductionHeure').getContext('2d');
    
    // Détruire le graphique existant
    if (charts.productionHeure) {
        charts.productionHeure.destroy();
    }
    
    // Agréger les données par heure
    const dataParHeure = {};
    donneesFiltrees.forEach(ligne => {
        const heure = ligne['Heure'] || '?';
        if (!dataParHeure[heure]) {
            dataParHeure[heure] = { bonne: 0, rebuts: 0 };
        }
        dataParHeure[heure].bonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
        dataParHeure[heure].rebuts += parseInt(ligne['Rebuts'] || 0);
    });
    
    // Trier les heures
    const heures = Object.keys(dataParHeure).sort((a, b) => {
        const numA = parseInt(a);
        const numB = parseInt(b);
        return numA - numB;
    });
    
    const prodBonne = heures.map(h => dataParHeure[h].bonne);
    const rebuts = heures.map(h => dataParHeure[h].rebuts);
    
    charts.productionHeure = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: heures,
            datasets: [
                {
                    label: 'Production Bonne',
                    data: prodBonne,
                    backgroundColor: 'rgba(75, 192, 192, 0.6)',
                    borderColor: 'rgba(75, 192, 192, 1)',
                    borderWidth: 2
                },
                {
                    label: 'Rebuts',
                    data: rebuts,
                    backgroundColor: 'rgba(255, 99, 132, 0.6)',
                    borderColor: 'rgba(255, 99, 132, 1)',
                    borderWidth: 2
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}

// Graphique: Répartition des arrêts
function afficherGraphiqueArrets() {
    const ctx = document.getElementById('chartArrets').getContext('2d');
    
    if (charts.arrets) {
        charts.arrets.destroy();
    }
    
    // Compter les arrêts par catégorie
    let countEquipement = 0;
    let countQualite = 0;
    let countOrganisation = 0;
    let countAutres = 0;
    
    donneesFiltrees.forEach(ligne => {
        if (ligne['Équipement'] || ligne['Équipement_x0020_Durée']) countEquipement++;
        if (ligne['Qualité'] || ligne['Qualité_x0020_Durée']) countQualite++;
        if (ligne['Organisation'] || ligne['Organisation_x0020_Durée']) countOrganisation++;
        if (ligne['Autres'] || ligne['Autres_x0020_Durée']) countAutres++;
    });
    
    charts.arrets = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['Équipement', 'Qualité', 'Organisation', 'Autres'],
            datasets: [{
                data: [countEquipement, countQualite, countOrganisation, countAutres],
                backgroundColor: [
                    'rgba(255, 159, 64, 0.6)',
                    'rgba(255, 99, 132, 0.6)',
                    'rgba(54, 162, 235, 0.6)',
                    'rgba(153, 102, 255, 0.6)'
                ],
                borderColor: [
                    'rgba(255, 159, 64, 1)',
                    'rgba(255, 99, 132, 1)',
                    'rgba(54, 162, 235, 1)',
                    'rgba(153, 102, 255, 1)'
                ],
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    position: 'bottom'
                }
            }
        }
    });
}

// Graphique: Production par équipe
function afficherGraphiqueProductionEquipe() {
    const ctx = document.getElementById('chartProductionEquipe').getContext('2d');
    
    if (charts.productionEquipe) {
        charts.productionEquipe.destroy();
    }
    
    // Agréger par équipe
    const dataParEquipe = {};
    donneesFiltrees.forEach(ligne => {
        const equipe = ligne['Équipe'] || 'Non définie';
        if (!dataParEquipe[equipe]) {
            dataParEquipe[equipe] = { bonne: 0, rebuts: 0 };
        }
        dataParEquipe[equipe].bonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
        dataParEquipe[equipe].rebuts += parseInt(ligne['Rebuts'] || 0);
    });
    
    const equipes = Object.keys(dataParEquipe);
    const prodBonne = equipes.map(e => dataParEquipe[e].bonne);
    const rebuts = equipes.map(e => dataParEquipe[e].rebuts);
    
    charts.productionEquipe = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: equipes,
            datasets: [
                {
                    label: 'Production Bonne',
                    data: prodBonne,
                    backgroundColor: 'rgba(75, 192, 192, 0.6)',
                    borderColor: 'rgba(75, 192, 192, 1)',
                    borderWidth: 2
                },
                {
                    label: 'Rebuts',
                    data: rebuts,
                    backgroundColor: 'rgba(255, 99, 132, 0.6)',
                    borderColor: 'rgba(255, 99, 132, 1)',
                    borderWidth: 2
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}

// Graphique: Durée des arrêts par catégorie
function afficherGraphiqueDureeArrets() {
    const ctx = document.getElementById('chartDureeArrets').getContext('2d');
    
    if (charts.dureeArrets) {
        charts.dureeArrets.destroy();
    }
    
    // Calculer la durée totale par catégorie
    let dureeEquipement = 0;
    let dureeQualite = 0;
    let dureeOrganisation = 0;
    let dureeAutres = 0;
    
    donneesFiltrees.forEach(ligne => {
        dureeEquipement += parseInt(ligne['Équipement Durée'] || ligne['Équipement_x0020_Durée'] || 0);
        dureeQualite += parseInt(ligne['Qualité Durée'] || ligne['Qualité_x0020_Durée'] || 0);
        dureeOrganisation += parseInt(ligne['Organisation Durée'] || ligne['Organisation_x0020_Durée'] || 0);
        dureeAutres += parseInt(ligne['Autres Durée'] || ligne['Autres_x0020_Durée'] || 0);
    });
    
    charts.dureeArrets = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ['Équipement', 'Qualité', 'Organisation', 'Autres'],
            datasets: [{
                label: 'Durée (minutes)',
                data: [dureeEquipement, dureeQualite, dureeOrganisation, dureeAutres],
                backgroundColor: [
                    'rgba(255, 159, 64, 0.6)',
                    'rgba(255, 99, 132, 0.6)',
                    'rgba(54, 162, 235, 0.6)',
                    'rgba(153, 102, 255, 0.6)'
                ],
                borderColor: [
                    'rgba(255, 159, 64, 1)',
                    'rgba(255, 99, 132, 1)',
                    'rgba(54, 162, 235, 1)',
                    'rgba(153, 102, 255, 1)'
                ],
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}

// Afficher le tableau détaillé
function afficherTableau() {
    const tbody = document.getElementById('dataTableBody');
    tbody.innerHTML = '';
    
    donneesFiltrees.forEach(ligne => {
        const row = tbody.insertRow();
        
        // Gérer les noms de colonnes avec ou sans espaces encodés
        const date = ligne['Date'] || '';
        const jour = ligne['Jour'] || '';
        const equipe = ligne['Équipe'] || '';
        const heure = ligne['Heure'] || '';
        const prodBonne = ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0;
        const rebuts = ligne['Rebuts'] || 0;
        const equipement = ligne['Équipement'] || '';
        const equipementDuree = ligne['Équipement Durée'] || ligne['Équipement_x0020_Durée'] || 0;
        const qualite = ligne['Qualité'] || '';
        const qualiteDuree = ligne['Qualité Durée'] || ligne['Qualité_x0020_Durée'] || 0;
        const organisation = ligne['Organisation'] || '';
        const organisationDuree = ligne['Organisation Durée'] || ligne['Organisation_x0020_Durée'] || 0;
        const autres = ligne['Autres'] || '';
        const autresDuree = ligne['Autres Durée'] || ligne['Autres_x0020_Durée'] || 0;
        const commentaire = ligne['Commentaire'] || '';
        
        row.innerHTML = `
            <td>${date}</td>
            <td>${jour}</td>
            <td>${equipe}</td>
            <td>${heure}</td>
            <td>${prodBonne}</td>
            <td>${rebuts}</td>
            <td>${equipement}</td>
            <td>${equipementDuree}</td>
            <td>${qualite}</td>
            <td>${qualiteDuree}</td>
            <td>${organisation}</td>
            <td>${organisationDuree}</td>
            <td>${autres}</td>
            <td>${autresDuree}</td>
            <td>${commentaire}</td>
        `;
    });
}

// Appliquer les filtres
function appliquerFiltres() {
    const filterDate = document.getElementById('filterDate').value;
    const filterEquipe = document.getElementById('filterEquipe').value;
    const filterJour = document.getElementById('filterJour').value;
    
    donneesFiltrees = donneesCompletes.filter(ligne => {
        let match = true;
        
        if (filterDate && ligne['Date'] !== filterDate) {
            match = false;
        }
        
        if (filterEquipe && ligne['Équipe'] !== filterEquipe) {
            match = false;
        }
        
        if (filterJour && ligne['Jour'] !== filterJour) {
            match = false;
        }
        
        return match;
    });
    
    console.log('Filtres appliqués:', donneesFiltrees.length, 'lignes');
    afficherDashboard();
}

// Réinitialiser les filtres
function reinitialiserFiltres() {
    document.getElementById('filterDate').value = '';
    document.getElementById('filterEquipe').value = '';
    document.getElementById('filterJour').value = '';
    
    donneesFiltrees = [...donneesCompletes];
    afficherDashboard();
}

// Exporter en CSV
function exporterCSV() {
    if (donneesFiltrees.length === 0) {
        alert('Aucune donnée à exporter');
        return;
    }
    
    // Créer l'en-tête CSV
    const headers = ['Date', 'Jour', 'Équipe', 'Heure', 'Prod Bonne', 'Rebuts', 
                    'Équipement', 'Équipement Durée', 'Qualité', 'Qualité Durée',
                    'Organisation', 'Organisation Durée', 'Autres', 'Autres Durée', 'Commentaire'];
    
    let csvContent = headers.join(';') + '\n';
    
    // Ajouter les données
    donneesFiltrees.forEach(ligne => {
        const row = [
            ligne['Date'] || '',
            ligne['Jour'] || '',
            ligne['Équipe'] || '',
            ligne['Heure'] || '',
            ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0,
            ligne['Rebuts'] || 0,
            ligne['Équipement'] || '',
            ligne['Équipement Durée'] || ligne['Équipement_x0020_Durée'] || 0,
            ligne['Qualité'] || '',
            ligne['Qualité Durée'] || ligne['Qualité_x0020_Durée'] || 0,
            ligne['Organisation'] || '',
            ligne['Organisation Durée'] || ligne['Organisation_x0020_Durée'] || 0,
            ligne['Autres'] || '',
            ligne['Autres Durée'] || ligne['Autres_x0020_Durée'] || 0,
            ligne['Commentaire'] || ''
        ];
        csvContent += row.join(';') + '\n';
    });
    
    // Télécharger le fichier
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', `HPH_CP17_Export_${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Initialisation
document.addEventListener('DOMContentLoaded', function() {
    console.log('✅ Dashboard HPH CP17 initialisé');
    chargerDonnees();
});
