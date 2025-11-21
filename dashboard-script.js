// URL du webhook Power Automate pour LIRE les données
const WEBHOOK_READ_URL = "https://defaultefdc051fe1144edaa7c5efec56da13.b1.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/68fb07af56e94845b714ce22d00b5f4c/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=hTGYXli6b3lv0VdHzQD8lKdAc07jyjerdzHloXKOV4c";

// Variables globales
let donneesCompletes = [];
let donneesFiltrees = [];
let charts = {};

// ✨ Fonction pour convertir les dates Excel (format numérique) en vraies dates
function convertirDateExcel(valeur) {
    // Si c'est déjà une date valide (string avec / ou -), la retourner
    if (typeof valeur === 'string' && (valeur.includes('/') || valeur.includes('-'))) {
        return valeur;
    }
    
    // ✨ Si c'est un STRING qui contient un nombre, le convertir en nombre
    let valeurNumerique = valeur;
    if (typeof valeur === 'string' && !isNaN(valeur) && valeur.trim() !== '') {
        valeurNumerique = parseFloat(valeur);
    }
    
    // Si c'est un nombre (date série Excel)
    if (typeof valeurNumerique === 'number' && valeurNumerique > 1000) {
        // Excel stocke les dates comme le nombre de jours depuis le 01/01/1900
        // Attention : Excel a un bug connu (compte 1900 comme année bissextile)
        const dateExcelEpoch = new Date(1899, 11, 30); // 30 décembre 1899
        const dateMs = dateExcelEpoch.getTime() + (valeurNumerique * 86400000); // 86400000 ms = 1 jour
        const dateObj = new Date(dateMs);
        
        // Retourner au format ISO (YYYY-MM-DD)
        const annee = dateObj.getFullYear();
        const mois = String(dateObj.getMonth() + 1).padStart(2, '0');
        const jour = String(dateObj.getDate()).padStart(2, '0');
        
        const resultat = `${annee}-${mois}-${jour}`;
        return resultat;
    }
    
    // Sinon, retourner tel quel
    return valeur;
}

// Objectifs par défaut (stockés dans localStorage)
let objectifs = {
    prodEquipe: 5000,
    tauxRebuts: 5,
    trs: 85
};

// Charger les objectifs depuis localStorage
function chargerObjectifs() {
    const stored = localStorage.getItem('objectifs_hph_cp17');
    if (stored) {
        objectifs = JSON.parse(stored);
        document.getElementById('objectifProdEquipe').value = objectifs.prodEquipe;
        document.getElementById('objectifTauxRebuts').value = objectifs.tauxRebuts;
        document.getElementById('objectifTRS').value = objectifs.trs || 85;
    }
}

// Enregistrer les objectifs
function appliquerObjectifs() {
    objectifs.prodEquipe = parseInt(document.getElementById('objectifProdEquipe').value) || 5000;
    objectifs.tauxRebuts = parseFloat(document.getElementById('objectifTauxRebuts').value) || 5;
    objectifs.trs = parseFloat(document.getElementById('objectifTRS').value) || 85;
    
    localStorage.setItem('objectifs_hph_cp17', JSON.stringify(objectifs));
    
    alert('✅ Objectifs enregistrés avec succès !');
    
    // Recalculer les KPIs avec les nouveaux objectifs
    if (donneesFiltrees.length > 0) {
        calculerKPIs();
        afficherGraphiques();
    }
}

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
        
        // Appliquer automatiquement le filtre sur l'équipe en cours
        appliquerFiltreEquipeParDefaut();
        
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
    
    // Comparer avec les objectifs
    compareAvecObjectifs(totalProduction, tauxRebuts, totalTempsArret);
}

// Comparer les résultats avec les objectifs
function compareAvecObjectifs(totalProduction, tauxRebuts, totalTempsArret) {
    const tauxRebutsNum = parseFloat(tauxRebuts);
    
    // Compter le nombre d'équipes différentes dans les données filtrées
    const equipesUniques = new Set();
    donneesFiltrees.forEach(ligne => {
        const equipe = ligne['Équipe'];
        if (equipe) {
            equipesUniques.add(equipe);
        }
    });
    const nombreEquipes = equipesUniques.size;
    
    // Calculer l'objectif total en fonction du nombre d'équipes
    const objectifTotal = objectifs.prodEquipe * nombreEquipes;
    
    // Production (par équipe)
    const kpiCardProd = document.getElementById('kpiCardProdTotal');
    const kpiObjectifProd = document.getElementById('kpiObjectifProd');
    const kpiStatusProd = document.getElementById('kpiStatusProd');
    
    if (nombreEquipes > 1) {
        kpiObjectifProd.textContent = `Objectif: ${objectifTotal.toLocaleString()} (${nombreEquipes} équipes × ${objectifs.prodEquipe.toLocaleString()})`;
    } else {
        kpiObjectifProd.textContent = `Objectif: ${objectifTotal.toLocaleString()} (${nombreEquipes} équipe)`;
    }
    
    if (totalProduction >= objectifTotal) {
        kpiStatusProd.textContent = '✅ Objectif atteint';
        kpiStatusProd.className = 'kpi-status atteint';
        kpiCardProd.className = 'kpi-card success';
    } else {
        const pourcentage = ((totalProduction / objectifTotal) * 100).toFixed(0);
        kpiStatusProd.textContent = `⚠️ ${pourcentage}% de l'objectif`;
        kpiStatusProd.className = 'kpi-status non-atteint';
        kpiCardProd.className = 'kpi-card warning';
    }
    
    // Taux de rebuts
    const kpiCardRebuts = document.getElementById('kpiCardTauxRebuts');
    const kpiObjectifRebuts = document.getElementById('kpiObjectifRebuts');
    const kpiStatusRebuts = document.getElementById('kpiStatusRebuts');
    
    kpiObjectifRebuts.textContent = `Objectif max: ${objectifs.tauxRebuts}%`;
    
    if (tauxRebutsNum <= objectifs.tauxRebuts) {
        kpiStatusRebuts.textContent = '✅ Objectif respecté';
        kpiStatusRebuts.className = 'kpi-status atteint';
        kpiCardRebuts.className = 'kpi-card success';
    } else {
        kpiStatusRebuts.textContent = '❌ Objectif dépassé';
        kpiStatusRebuts.className = 'kpi-status non-atteint';
        kpiCardRebuts.className = 'kpi-card danger';
    }
}

// Afficher les graphiques
function afficherGraphiques() {
    afficherGraphiqueProductionHeure();
    afficherGraphiqueTRS();
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
            dataParHeure[heure] = { bonne: 0 };
        }
        dataParHeure[heure].bonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
    });
    
    // Trier les heures
    const heures = Object.keys(dataParHeure).sort((a, b) => {
        const numA = parseInt(a);
        const numB = parseInt(b);
        return numA - numB;
    });
    
    const prodBonne = heures.map(h => dataParHeure[h].bonne);
    
    // Calculer l'objectif par heure (objectif équipe / 8 heures)
    const objectifParHeure = objectifs.prodEquipe / 8;
    const objectifData = heures.map(() => objectifParHeure);
    
    // Code couleur : vert si >= objectif, rouge si < objectif
    const backgroundColor = prodBonne.map(val => 
        val >= objectifParHeure ? 'rgba(75, 192, 75, 0.6)' : 'rgba(255, 99, 99, 0.6)'
    );
    
    const borderColor = prodBonne.map(val => 
        val >= objectifParHeure ? 'rgba(75, 192, 75, 1)' : 'rgba(255, 99, 99, 1)'
    );
    
    charts.productionHeure = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: heures,
            datasets: [
                {
                    label: 'Production Bonne',
                    data: prodBonne,
                    backgroundColor: backgroundColor,
                    borderColor: borderColor,
                    borderWidth: 2,
                    order: 2
                },
                {
                    label: 'Objectif/Heure',
                    data: objectifData,
                    type: 'line',
                    borderColor: 'rgba(255, 193, 7, 1)',
                    backgroundColor: 'rgba(255, 193, 7, 0.1)',
                    borderWidth: 3,
                    borderDash: [10, 5],
                    pointRadius: 5,
                    pointBackgroundColor: 'rgba(255, 193, 7, 1)',
                    fill: false,
                    order: 1,
                    tension: 0
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Quantité'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Heure'
                    }
                }
            },
            plugins: {
                legend: {
                    display: true,
                    position: 'top'
                },
                tooltip: {
                    callbacks: {
                        afterLabel: function(context) {
                            if (context.datasetIndex === 0) {
                                const value = context.parsed.y;
                                const objectif = objectifParHeure;
                                const ecart = value - objectif;
                                if (ecart >= 0) {
                                    return `✅ +${ecart.toFixed(0)} vs objectif`;
                                } else {
                                    return `❌ ${ecart.toFixed(0)} vs objectif`;
                                }
                            }
                            return '';
                        }
                    }
                }
            }
        }
    });
}

// Graphique: TRS par Heure
function afficherGraphiqueTRS() {
    const ctx = document.getElementById('chartTRS').getContext('2d');
    
    if (charts.trs) {
        charts.trs.destroy();
    }
    
    // Calculer le TRS par heure
    const dataParHeure = {};
    donneesFiltrees.forEach(ligne => {
        const heure = ligne['Heure'] || '?';
        if (!dataParHeure[heure]) {
            dataParHeure[heure] = { bonne: 0 };
        }
        dataParHeure[heure].bonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
    });
    
    // Trier les heures
    const heures = Object.keys(dataParHeure).sort((a, b) => {
        const numA = parseInt(a);
        const numB = parseInt(b);
        return numA - numB;
    });
    
    // Calculer le TRS pour chaque heure : (Production Bonne / Objectif par heure) × 100
    const objectifParHeure = objectifs.prodEquipe / 8;
    const trsData = heures.map(h => {
        const prodBonne = dataParHeure[h].bonne;
        const trs = (prodBonne / objectifParHeure) * 100;
        return Math.min(trs, 150); // Limiter à 150% pour l'affichage
    });
    
    // Ligne d'objectif TRS
    const objectifTRSData = heures.map(() => objectifs.trs);
    
    // Code couleur : vert si >= objectif TRS, rouge si < objectif TRS
    const backgroundColor = trsData.map(val => 
        val >= objectifs.trs ? 'rgba(75, 192, 192, 0.6)' : 'rgba(255, 99, 132, 0.6)'
    );
    
    const borderColor = trsData.map(val => 
        val >= objectifs.trs ? 'rgba(75, 192, 192, 1)' : 'rgba(255, 99, 132, 1)'
    );
    
    charts.trs = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: heures,
            datasets: [
                {
                    label: 'TRS Réalisé (%)',
                    data: trsData,
                    backgroundColor: backgroundColor,
                    borderColor: borderColor,
                    borderWidth: 2,
                    order: 2
                },
                {
                    label: 'Objectif TRS (%)',
                    data: objectifTRSData,
                    type: 'line',
                    borderColor: 'rgba(255, 193, 7, 1)',
                    backgroundColor: 'rgba(255, 193, 7, 0.1)',
                    borderWidth: 3,
                    borderDash: [10, 5],
                    pointRadius: 5,
                    pointBackgroundColor: 'rgba(255, 193, 7, 1)',
                    fill: false,
                    order: 1,
                    tension: 0
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            scales: {
                y: {
                    beginAtZero: true,
                    max: 150,
                    title: {
                        display: true,
                        text: 'TRS (%)'
                    },
                    ticks: {
                        callback: function(value) {
                            return value + '%';
                        }
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Heure'
                    }
                }
            },
            plugins: {
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return context.dataset.label + ': ' + context.parsed.y.toFixed(1) + '%';
                        },
                        afterLabel: function(context) {
                            if (context.datasetIndex === 0) {
                                const value = context.parsed.y;
                                const objectif = objectifs.trs;
                                const ecart = value - objectif;
                                if (ecart >= 0) {
                                    return `✅ +${ecart.toFixed(1)}% vs objectif`;
                                } else {
                                    return `❌ ${ecart.toFixed(1)}% vs objectif`;
                                }
                            }
                            return '';
                        }
                    }
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
        let date = ligne['Date'] || '';
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
        
        // ✨ D'ABORD convertir la date Excel si nécessaire
        date = convertirDateExcel(date);
        
        // PUIS formater la date au format français (JJ/MM/AAAA)
        if (date) {
            try {
                const dateObj = new Date(date);
                if (!isNaN(dateObj.getTime())) {
                    const jourDate = String(dateObj.getDate()).padStart(2, '0');
                    const mois = String(dateObj.getMonth() + 1).padStart(2, '0');
                    const annee = dateObj.getFullYear();
                    date = `${jourDate}/${mois}/${annee}`;
                }
            } catch (e) {
                console.log('Erreur de formatage de date:', e);
            }
        }
        
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
    const filterDateDebut = document.getElementById('filterDateDebut').value;
    const filterDateFin = document.getElementById('filterDateFin').value;
    const filterEquipe = document.getElementById('filterEquipe').value;
    
    console.log('Filtres demandés:', { filterDateDebut, filterDateFin, filterEquipe });
    
    donneesFiltrees = donneesCompletes.filter(ligne => {
        let match = true;
        
        // Filtre de date (plage ou date unique)
        if (filterDateDebut || filterDateFin) {
            // ✨ Convertir la date Excel avant de comparer
            let ligneDate = convertirDateExcel(ligne['Date'] || '');
            
            // Convertir la date en format ISO pour comparaison
            let ligneDateISO = ligneDate;
            
            // Si la date est au format français (21/11/2025), la convertir en ISO
            if (ligneDate.includes('/')) {
                const parts = ligneDate.split('/');
                if (parts.length === 3) {
                    ligneDateISO = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
                }
            }
            
            // Vérifier si la date est dans la plage
            if (filterDateDebut && ligneDateISO < filterDateDebut) {
                match = false;
            }
            if (filterDateFin && ligneDateISO > filterDateFin) {
                match = false;
            }
        }
        
        // Filtre d'équipe
        if (filterEquipe && ligne['Équipe'] !== filterEquipe) {
            match = false;
        }
        
        return match;
    });
    
    console.log('Résultats filtrés:', donneesFiltrees.length, 'lignes sur', donneesCompletes.length);
    
    if (donneesFiltrees.length === 0) {
        alert('⚠️ Aucune donnée ne correspond aux filtres sélectionnés.');
        return;
    }
    
    afficherDashboard();
    alert(`✅ Filtres appliqués : ${donneesFiltrees.length} ligne(s) affichée(s)`);
}

// Réinitialiser les filtres
function reinitialiserFiltres() {
    document.getElementById('filterDateDebut').value = '';
    document.getElementById('filterDateFin').value = '';
    document.getElementById('filterEquipe').value = '';
    
    donneesFiltrees = [...donneesCompletes];
    
    // Réappliquer le filtre par défaut sur l'équipe en cours
    appliquerFiltreEquipeParDefaut();
}

// Déterminer l'équipe en cours selon l'heure
function determinerEquipeEnCours() {
    const heure = new Date().getHours();
    let equipe = '';
    
    if (heure >= 6 && heure < 14) {
        equipe = 'Matin';
    } else if (heure >= 14 && heure < 22) {
        equipe = 'Soir';
    } else {
        equipe = 'Nuit';
    }
    
    return equipe;
}

// Appliquer le filtre par défaut sur l'équipe en cours
function appliquerFiltreEquipeParDefaut() {
    const equipeEnCours = determinerEquipeEnCours();
    
    console.log('🕐 Équipe en cours détectée:', equipeEnCours);
    
    // Pré-remplir le filtre avec l'équipe en cours
    document.getElementById('filterEquipe').value = equipeEnCours;
    
    // Filtrer les données sur l'équipe en cours
    donneesFiltrees = donneesCompletes.filter(ligne => {
        return ligne['Équipe'] === equipeEnCours;
    });
    
    console.log('📊 Données filtrées sur équipe', equipeEnCours, ':', donneesFiltrees.length, 'lignes');
    
    // Afficher le dashboard avec les données filtrées
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
        let date = ligne['Date'] || '';
        
        // ✨ D'ABORD convertir la date Excel si nécessaire
        date = convertirDateExcel(date);
        
        // PUIS formater la date au format français
        if (date) {
            try {
                const dateObj = new Date(date);
                if (!isNaN(dateObj.getTime())) {
                    const jour = String(dateObj.getDate()).padStart(2, '0');
                    const mois = String(dateObj.getMonth() + 1).padStart(2, '0');
                    const annee = dateObj.getFullYear();
                    date = `${jour}/${mois}/${annee}`;
                }
            } catch (e) {
                console.log('Erreur de formatage de date:', e);
            }
        }
        
        const row = [
            date,
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
    chargerObjectifs();
    chargerDonnees();
});
