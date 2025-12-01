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
    trs: 85,
    cadence: 15
};

// Charger les objectifs depuis localStorage
function chargerObjectifs() {
    const stored = localStorage.getItem('objectifs_hph_cp17');
    if (stored) {
        objectifs = JSON.parse(stored);
        document.getElementById('objectifProdEquipe').value = objectifs.prodEquipe;
        document.getElementById('objectifTauxRebuts').value = objectifs.tauxRebuts;
        document.getElementById('objectifTRS').value = objectifs.trs || 85;
        document.getElementById('cadenceInstantanee').value = objectifs.cadence || 15;
    }
}

// Enregistrer les objectifs
function appliquerObjectifs() {
    objectifs.prodEquipe = parseInt(document.getElementById('objectifProdEquipe').value) || 5000;
    objectifs.tauxRebuts = parseFloat(document.getElementById('objectifTauxRebuts').value) || 5;
    objectifs.trs = parseFloat(document.getElementById('objectifTRS').value) || 85;
    objectifs.cadence = parseFloat(document.getElementById('cadenceInstantanee').value) || 15;
    
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
    
    // Compter le nombre d'heures dans les données filtrées
    const heuresUniques = new Set();
    
    donneesFiltrees.forEach(ligne => {
        const heure = ligne['Heure'] || '?';
        heuresUniques.add(heure);
        
        const prodBonne = parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
        const rebuts = parseInt(ligne['Rebuts'] || 0);
        
        totalProdBonne += prodBonne;
        totalRebuts += rebuts;
        
        // Calculer le temps d'arrêt total
        const equipDuree = parseInt(ligne['Équipement Durée'] || ligne['Équipement_x0020_Durée'] || 0);
        const qualiteDuree = parseInt(ligne['Qualité Durée'] || ligne['Qualité_x0020_Durée'] || 0);
        const orgDuree = parseInt(ligne['Organisation Durée'] || ligne['Organisation_x0020_Durée'] || 0);
        const autresDuree = parseInt(ligne['Autres Durée'] || ligne['Autres_x0020_Durée'] || 0);
        
        totalTempsArret += equipDuree + qualiteDuree + orgDuree + autresDuree;
    });
    
    const totalProduction = totalProdBonne + totalRebuts;
    const tauxRebuts = totalProduction > 0 ? ((totalRebuts / totalProduction) * 100).toFixed(2) : 0;
    
    // Calculer le TRS Global : (Prod Totale / (Cadence × 60 × Nb heures)) × 100
    const cadence = objectifs.cadence;
    const nbHeures = heuresUniques.size;
    const prodMaxGlobale = cadence * 60 * nbHeures;
    const trsGlobal = prodMaxGlobale > 0 ? ((totalProduction / prodMaxGlobale) * 100).toFixed(1) : 0;
    
    // Afficher les KPIs
    document.getElementById('kpiProdTotal').textContent = totalProduction.toLocaleString();
    document.getElementById('kpiProdBonne').textContent = totalProdBonne.toLocaleString();
    document.getElementById('kpiRebuts').textContent = totalRebuts.toLocaleString();
    document.getElementById('kpiTauxRebuts').textContent = tauxRebuts + '%';
    document.getElementById('kpiTempsArret').textContent = totalTempsArret + ' min';
    document.getElementById('kpiTRSGlobal').textContent = trsGlobal + '%';
    
    // Comparer avec les objectifs
    compareAvecObjectifs(totalProduction, tauxRebuts, totalTempsArret, parseFloat(trsGlobal));
}

// Comparer les résultats avec les objectifs
function compareAvecObjectifs(totalProduction, tauxRebuts, totalTempsArret, trsGlobal) {
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
    
    // TRS Global
    const kpiCardTRSGlobal = document.getElementById('kpiCardTRSGlobal');
    const kpiObjectifTRSGlobal = document.getElementById('kpiObjectifTRSGlobal');
    const kpiStatusTRSGlobal = document.getElementById('kpiStatusTRSGlobal');
    
    kpiObjectifTRSGlobal.textContent = `Objectif: ${objectifs.trs}%`;
    
    if (trsGlobal >= objectifs.trs) {
        kpiStatusTRSGlobal.textContent = '✅ Objectif atteint';
        kpiStatusTRSGlobal.className = 'kpi-status atteint';
        kpiCardTRSGlobal.className = 'kpi-card success';
    } else {
        kpiStatusTRSGlobal.textContent = `⚠️ ${trsGlobal}% (objectif: ${objectifs.trs}%)`;
        kpiStatusTRSGlobal.className = 'kpi-status non-atteint';
        kpiCardTRSGlobal.className = 'kpi-card warning';
    }
}

// Afficher les graphiques
function afficherGraphiques() {
    afficherGraphiqueProductionHeure();
    afficherGraphiqueTRS();
    afficherGraphiquePerformance();
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
                    position: 'bottom'
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
    
    // Calculer le TRS par heure avec la nouvelle formule
    // TRS = (Prod Totale / Prod Théorique) × 100
    // Prod Théorique = Cadence (pièces/min) × (60 min - temps d'arrêts)
    const dataParHeure = {};
    donneesFiltrees.forEach(ligne => {
        const heure = ligne['Heure'] || '?';
        if (!dataParHeure[heure]) {
            dataParHeure[heure] = { 
                bonne: 0, 
                rebuts: 0, 
                arrets: 0 
            };
        }
        
        // Production totale (bonne + rebuts)
        dataParHeure[heure].bonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
        dataParHeure[heure].rebuts += parseInt(ligne['Rebuts'] || 0);
        
        // Temps d'arrêts (en minutes)
        const equipDuree = parseInt(ligne['Équipement Durée'] || ligne['Équipement_x0020_Durée'] || 0);
        const qualiteDuree = parseInt(ligne['Qualité Durée'] || ligne['Qualité_x0020_Durée'] || 0);
        const orgDuree = parseInt(ligne['Organisation Durée'] || ligne['Organisation_x0020_Durée'] || 0);
        const autresDuree = parseInt(ligne['Autres Durée'] || ligne['Autres_x0020_Durée'] || 0);
        
        dataParHeure[heure].arrets += equipDuree + qualiteDuree + orgDuree + autresDuree;
    });
    
    // Trier les heures
    const heures = Object.keys(dataParHeure).sort((a, b) => {
        const numA = parseInt(a);
        const numB = parseInt(b);
        return numA - numB;
    });
    
    // Calculer le TRS pour chaque heure
    // NOUVELLE FORMULE: TRS = (Prod Totale / (Cadence × 60)) × 100
    const cadence = objectifs.cadence; // pièces/minute
    const trsData = heures.map(h => {
        const prodTotale = dataParHeure[h].bonne + dataParHeure[h].rebuts;
        const prodMaxMachine = cadence * 60; // Production max machine en 60 minutes
        
        const trs = prodMaxMachine > 0 ? (prodTotale / prodMaxMachine) * 100 : 0;
        
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
            aspectRatio: 2.2,
            layout: {
                padding: {
                    top: 40,
                    bottom: 10,
                    left: 10,
                    right: 10
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    max: 170,
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
                legend: {
                    position: 'bottom',
                    labels: {
                        padding: 15,
                        font: {
                            size: 12
                        }
                    }
                },
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
            },
            animation: {
                onComplete: function() {
                    const chart = this;
                    const ctx = chart.ctx;
                    
                    ctx.font = 'bold 11px Arial';
                    ctx.fillStyle = '#333';
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'bottom';
                    
                    chart.data.datasets[0].data.forEach((value, index) => {
                        const meta = chart.getDatasetMeta(0);
                        const bar = meta.data[index];
                        
                        if (bar) {
                            const text = value.toFixed(1) + '%';
                            ctx.fillText(text, bar.x, bar.y - 5);
                        }
                    });
                }
            }
        }
    });
}

// Graphique: Taux de Performance par Heure
function afficherGraphiquePerformance() {
    const ctx = document.getElementById('chartPerformance').getContext('2d');
    
    if (charts.performance) {
        charts.performance.destroy();
    }
    
    // Calculer le Taux de Performance par heure
    // Performance = (Prod Totale / Prod Théorique) × 100
    // Prod Théorique = Cadence (pièces/min) × (60 min - temps d'arrêts)
    const dataParHeure = {};
    donneesFiltrees.forEach(ligne => {
        const heure = ligne['Heure'] || '?';
        if (!dataParHeure[heure]) {
            dataParHeure[heure] = { 
                bonne: 0, 
                rebuts: 0, 
                arrets: 0 
            };
        }
        
        // Production totale (bonne + rebuts)
        dataParHeure[heure].bonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
        dataParHeure[heure].rebuts += parseInt(ligne['Rebuts'] || 0);
        
        // Temps d'arrêts (en minutes)
        const equipDuree = parseInt(ligne['Équipement Durée'] || ligne['Équipement_x0020_Durée'] || 0);
        const qualiteDuree = parseInt(ligne['Qualité Durée'] || ligne['Qualité_x0020_Durée'] || 0);
        const orgDuree = parseInt(ligne['Organisation Durée'] || ligne['Organisation_x0020_Durée'] || 0);
        const autresDuree = parseInt(ligne['Autres Durée'] || ligne['Autres_x0020_Durée'] || 0);
        
        dataParHeure[heure].arrets += equipDuree + qualiteDuree + orgDuree + autresDuree;
    });
    
    // Trier les heures
    const heures = Object.keys(dataParHeure).sort((a, b) => {
        const numA = parseInt(a);
        const numB = parseInt(b);
        return numA - numB;
    });
    
    // Calculer la Performance pour chaque heure
    const cadence = objectifs.cadence; // pièces/minute
    const performanceData = heures.map(h => {
        const prodTotale = dataParHeure[h].bonne + dataParHeure[h].rebuts;
        const tempsArrets = dataParHeure[h].arrets; // en minutes
        const tempsProductif = 60 - tempsArrets; // 60 min - arrêts
        
        if (tempsProductif <= 0) {
            return 0; // Pas de production possible si arrêts >= 60 min
        }
        
        const prodTheorique = cadence * tempsProductif;
        const performance = prodTheorique > 0 ? (prodTotale / prodTheorique) * 100 : 0;
        
        return Math.min(performance, 150); // Limiter à 150% pour l'affichage
    });
    
    // Ligne d'objectif (100% = performance attendue)
    const objectifPerformanceData = heures.map(() => 100);
    
    // Code couleur : vert si >= 100%, orange si >= 85%, rouge sinon
    const backgroundColor = performanceData.map(val => {
        if (val >= 100) return 'rgba(75, 192, 75, 0.6)';
        if (val >= 85) return 'rgba(255, 193, 7, 0.6)';
        return 'rgba(255, 99, 132, 0.6)';
    });
    
    const borderColor = performanceData.map(val => {
        if (val >= 100) return 'rgba(75, 192, 75, 1)';
        if (val >= 85) return 'rgba(255, 193, 7, 1)';
        return 'rgba(255, 99, 132, 1)';
    });
    
    charts.performance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: heures,
            datasets: [
                {
                    label: 'Performance Réalisée (%)',
                    data: performanceData,
                    backgroundColor: backgroundColor,
                    borderColor: borderColor,
                    borderWidth: 2,
                    order: 2
                },
                {
                    label: 'Objectif Performance (100%)',
                    data: objectifPerformanceData,
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
            aspectRatio: 2.2,
            layout: {
                padding: {
                    top: 40,
                    bottom: 10,
                    left: 10,
                    right: 10
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    max: 170,
                    title: {
                        display: true,
                        text: 'Performance (%)'
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
                legend: {
                    position: 'bottom',
                    labels: {
                        padding: 15,
                        font: {
                            size: 12
                        }
                    }
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return context.dataset.label + ': ' + context.parsed.y.toFixed(1) + '%';
                        },
                        afterLabel: function(context) {
                            if (context.datasetIndex === 0) {
                                const value = context.parsed.y;
                                const heure = context.label;
                                const data = dataParHeure[heure];
                                const prodTotale = data.bonne + data.rebuts;
                                const tempsProductif = 60 - data.arrets;
                                const prodTheorique = cadence * tempsProductif;
                                
                                return [
                                    `Production: ${prodTotale} pcs`,
                                    `Prod Théorique: ${prodTheorique.toFixed(0)} pcs`,
                                    `Temps productif: ${tempsProductif} min`
                                ];
                            }
                            return '';
                        }
                    }
                }
            },
            animation: {
                onComplete: function() {
                    const chart = this;
                    const ctx = chart.ctx;
                    
                    ctx.font = 'bold 11px Arial';
                    ctx.fillStyle = '#333';
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'bottom';
                    
                    chart.data.datasets[0].data.forEach((value, index) => {
                        const meta = chart.getDatasetMeta(0);
                        const bar = meta.data[index];
                        
                        if (bar) {
                            const text = value.toFixed(1) + '%';
                            ctx.fillText(text, bar.x, bar.y - 5);
                        }
                    });
                }
            }
        }
    });
}

// Graphique: Temps non justifié par heure
function afficherGraphiqueProductionEquipe() {
    const ctx = document.getElementById('chartProductionEquipe').getContext('2d');
    
    if (charts.productionEquipe) {
        charts.productionEquipe.destroy();
    }
    
    // Calculer le temps non justifié pour chaque heure
    const dataParHeure = {};
    
    donneesFiltrees.forEach(ligne => {
        const heure = ligne['Heure'] || '?';
        
        if (!dataParHeure[heure]) {
            dataParHeure[heure] = {
                prodBonne: 0,
                rebuts: 0,
                arrets: 0
            };
        }
        
        dataParHeure[heure].prodBonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
        dataParHeure[heure].rebuts += parseInt(ligne['Rebuts'] || 0);
        
        // Temps d'arrêts
        const equipDuree = parseInt(ligne['Équipement Durée'] || ligne['Équipement_x0020_Durée'] || 0);
        const qualiteDuree = parseInt(ligne['Qualité Durée'] || ligne['Qualité_x0020_Durée'] || 0);
        const orgDuree = parseInt(ligne['Organisation Durée'] || ligne['Organisation_x0020_Durée'] || 0);
        const autresDuree = parseInt(ligne['Autres Durée'] || ligne['Autres_x0020_Durée'] || 0);
        
        dataParHeure[heure].arrets += equipDuree + qualiteDuree + orgDuree + autresDuree;
    });
    
    // Trier les heures
    const heures = Object.keys(dataParHeure).sort((a, b) => {
        const numA = parseInt(a);
        const numB = parseInt(b);
        return numA - numB;
    });
    
    // Calculer le temps non justifié pour chaque heure
    const cadence = objectifs.cadence; // pièces/minute
    const tempsNonJustifieData = heures.map(h => {
        const data = dataParHeure[h];
        const prodTotale = data.prodBonne + data.rebuts;
        const prodAttendue = cadence * 60; // Production attendue en 60 minutes
        
        if (prodTotale < prodAttendue) {
            const ecartPieces = prodAttendue - prodTotale;
            const tempsEquivalentManquant = ecartPieces / cadence; // en minutes
            const tempsNonJustifie = Math.max(0, tempsEquivalentManquant - data.arrets);
            
            return tempsNonJustifie;
        }
        
        return 0; // Pas de temps non justifié si production >= attendue
    });
    
    // Code couleur : rouge si > 5 min, orange si > 0, vert si 0
    const backgroundColor = tempsNonJustifieData.map(val => {
        if (val === 0) return 'rgba(75, 192, 75, 0.6)';
        if (val > 5) return 'rgba(255, 99, 99, 0.6)';
        return 'rgba(255, 193, 7, 0.6)';
    });
    
    const borderColor = tempsNonJustifieData.map(val => {
        if (val === 0) return 'rgba(75, 192, 75, 1)';
        if (val > 5) return 'rgba(255, 99, 99, 1)';
        return 'rgba(255, 193, 7, 1)';
    });
    
    charts.productionEquipe = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: heures,
            datasets: [
                {
                    label: 'Temps Non Justifié (min)',
                    data: tempsNonJustifieData,
                    backgroundColor: backgroundColor,
                    borderColor: borderColor,
                    borderWidth: 2
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
                        text: 'Minutes'
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
                    position: 'bottom'
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const value = context.parsed.y;
                            return `Temps non justifié: ${value.toFixed(1)} min`;
                        },
                        afterLabel: function(context) {
                            const heure = context.label;
                            const data = dataParHeure[heure];
                            const prodTotale = data.prodBonne + data.rebuts;
                            const prodAttendue = cadence * 60;
                            const ecart = prodAttendue - prodTotale;
                            
                            return [
                                `Production: ${prodTotale} pcs`,
                                `Attendue: ${prodAttendue} pcs`,
                                `Écart: ${ecart} pcs`,
                                `Arrêts justifiés: ${data.arrets} min`
                            ];
                        }
                    }
                }
            },
            animation: {
                onComplete: function() {
                    const chart = this;
                    const ctx = chart.ctx;
                    
                    ctx.font = 'bold 11px Arial';
                    ctx.fillStyle = '#333';
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'bottom';
                    
                    chart.data.datasets[0].data.forEach((value, index) => {
                        const meta = chart.getDatasetMeta(0);
                        const bar = meta.data[index];
                        
                        if (bar && value > 0) {
                            const text = value.toFixed(1) + ' min';
                            ctx.fillText(text, bar.x, bar.y - 5);
                        }
                    });
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
        const reference = ligne['Référence'] || ligne['Reference'] || '';
        const trigramme = ligne['Jour'] || ''; // Le trigramme est stocké dans la colonne "Jour"
        const equipe = ligne['Équipe'] || ligne['Equipe'] || ''; // Avec ou sans accent
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
            <td>${reference}</td>
            <td>${trigramme}</td>
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
    const dateAujourdhui = new Date().toISOString().split('T')[0]; // Format: 2025-11-24
    
    console.log('🕐 Équipe en cours détectée:', equipeEnCours);
    console.log('📅 Date du jour:', dateAujourdhui);
    console.log('📊 Nombre total de données:', donneesCompletes.length);
    
    // Debug: Afficher un exemple de ligne pour voir la structure
    if (donneesCompletes.length > 0) {
        console.log('📋 Exemple de ligne:', donneesCompletes[0]);
    }
    
    // Pré-remplir le filtre avec l'équipe en cours
    document.getElementById('filterEquipe').value = equipeEnCours;
    
    // Filtrer les données sur l'équipe en cours ET la date du jour
    donneesFiltrees = donneesCompletes.filter(ligne => {
        const equipe = ligne['Équipe'];
        let ligneDate = convertirDateExcel(ligne['Date'] || '');
        
        // Convertir la date en format ISO pour comparaison
        let ligneDateISO = ligneDate;
        if (ligneDate.includes('/')) {
            const parts = ligneDate.split('/');
            if (parts.length === 3) {
                ligneDateISO = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
            }
        }
        
        const equipeMatch = equipe === equipeEnCours;
        const dateMatch = ligneDateISO === dateAujourdhui;
        
        console.log('🔍 Ligne:', { equipe, date: ligneDateISO, equipeMatch, dateMatch });
        
        return equipeMatch && dateMatch;
    });
    
    console.log('📊 Données filtrées (équipe', equipeEnCours, '+ date', dateAujourdhui + '):', donneesFiltrees.length, 'lignes');
    
    // Si aucune donnée pour l'équipe en cours + date du jour, chercher la dernière date avec des données
    if (donneesFiltrees.length === 0) {
        console.log('⚠️ Aucune donnée pour l\'équipe', equipeEnCours, 'à la date', dateAujourdhui);
        console.log('🔍 Recherche de la dernière date avec des données...');
        
        // D'abord essayer de trouver la dernière date pour l'équipe en cours
        let donneesEquipeEnCours = donneesCompletes.filter(ligne => ligne['Équipe'] === equipeEnCours);
        
        if (donneesEquipeEnCours.length > 0) {
            // Trouver la date la plus récente pour cette équipe
            let datesUniques = new Set();
            donneesEquipeEnCours.forEach(ligne => {
                let ligneDate = convertirDateExcel(ligne['Date'] || '');
                if (ligneDate.includes('/')) {
                    const parts = ligneDate.split('/');
                    if (parts.length === 3) {
                        ligneDate = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
                    }
                }
                datesUniques.add(ligneDate);
            });
            
            const datesTriees = Array.from(datesUniques).sort().reverse(); // Plus récente en premier
            const derniereDateEquipe = datesTriees[0];
            
            console.log('✅ Dernière date pour équipe', equipeEnCours + ':', derniereDateEquipe);
            
            donneesFiltrees = donneesEquipeEnCours.filter(ligne => {
                let ligneDate = convertirDateExcel(ligne['Date'] || '');
                if (ligneDate.includes('/')) {
                    const parts = ligneDate.split('/');
                    if (parts.length === 3) {
                        ligneDate = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
                    }
                }
                return ligneDate === derniereDateEquipe;
            });
        } else {
            // Si aucune donnée pour l'équipe en cours, chercher la dernière équipe
            console.log('⚠️ Aucune donnée pour l\'équipe', equipeEnCours, '- Recherche de la dernière équipe avec des données');
            
            const ordreEquipes = ['Nuit', 'Soir', 'Matin'];
            const indexEquipeEnCours = ordreEquipes.indexOf(equipeEnCours);
            
            let equipeTrouvee = null;
            for (let i = 1; i < ordreEquipes.length; i++) {
                const index = (indexEquipeEnCours + i) % ordreEquipes.length;
                const equipeTest = ordreEquipes[index];
                
                const donneesTest = donneesCompletes.filter(ligne => ligne['Équipe'] === equipeTest);
                if (donneesTest.length > 0) {
                    equipeTrouvee = equipeTest;
                    
                    // Trouver la dernière date pour cette équipe
                    let datesUniques = new Set();
                    donneesTest.forEach(ligne => {
                        let ligneDate = convertirDateExcel(ligne['Date'] || '');
                        if (ligneDate.includes('/')) {
                            const parts = ligneDate.split('/');
                            if (parts.length === 3) {
                                ligneDate = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
                            }
                        }
                        datesUniques.add(ligneDate);
                    });
                    
                    const datesTriees = Array.from(datesUniques).sort().reverse();
                    const derniereDate = datesTriees[0];
                    
                    donneesFiltrees = donneesTest.filter(ligne => {
                        let ligneDate = convertirDateExcel(ligne['Date'] || '');
                        if (ligneDate.includes('/')) {
                            const parts = ligneDate.split('/');
                            if (parts.length === 3) {
                                ligneDate = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
                            }
                        }
                        return ligneDate === derniereDate;
                    });
                    
                    console.log('✅ Dernière équipe avec des données:', equipeTrouvee, 'à la date', derniereDate);
                    document.getElementById('filterEquipe').value = equipeTrouvee;
                    break;
                }
            }
            
            if (!equipeTrouvee) {
                console.log('⚠️ Aucune équipe n\'a de données - Affichage de toutes les données');
                donneesFiltrees = [...donneesCompletes];
                document.getElementById('filterEquipe').value = '';
            }
        }
    }
    
    console.log('📊 Résultat final:', donneesFiltrees.length, 'lignes affichées');
    
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
    const headers = ['Date', 'Référence', 'Trigramme', 'Équipe', 'Heure', 'Prod Bonne', 'Rebuts', 
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
            ligne['Référence'] || ligne['Reference'] || '',
            ligne['Jour'] || '', // Le trigramme est dans la colonne Jour
            ligne['Équipe'] || ligne['Equipe'] || '',
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

// Exporter le dashboard en PDF
async function exporterPDF() {
    // Cacher les éléments non nécessaires pour l'impression
    const headerActions = document.querySelector('.header-actions');
    const tableControls = document.querySelector('.table-controls');
    const filterSection = document.querySelector('.filters-section');
    const objectifsSection = document.querySelector('.objectifs-section');
    const loadingMessage = document.getElementById('loadingMessage');
    const errorMessage = document.getElementById('errorMessage');
    
    // Sauvegarder les états d'affichage
    const elementsToHide = [
        headerActions, 
        tableControls, 
        filterSection, 
        objectifsSection, 
        loadingMessage, 
        errorMessage
    ];
    
    const originalDisplayStates = elementsToHide.map(el => el ? el.style.display : '');
    
    // Cacher temporairement
    elementsToHide.forEach(el => {
        if (el) el.style.display = 'none';
    });
    
    // Forcer les couleurs vives sur les KPI cards
    const kpiCards = document.querySelectorAll('.kpi-card');
    const originalBackgrounds = [];
    kpiCards.forEach((card, index) => {
        originalBackgrounds[index] = card.style.background;
        // Couleurs solides pour meilleur rendu PDF
        if (index < 5) {
            card.style.background = '#667eea';
        } else {
            card.style.background = '#ff9800';
        }
        card.style.setProperty('background', card.style.background, 'important');
    });
    
    // Ajouter un titre personnalisé pour l'impression
    const titre = document.createElement('div');
    titre.id = 'pdf-titre-temp';
    titre.style.cssText = 'text-align: center; margin-bottom: 20px; font-size: 0.9em; color: #666;';
    titre.innerHTML = `Dashboard généré le ${new Date().toLocaleDateString('fr-FR')} à ${new Date().toLocaleTimeString('fr-FR')}`;
    document.querySelector('.container').insertBefore(titre, document.querySelector('.container').firstChild);
    
    // Message à l'utilisateur
    const originalTitle = document.title;
    document.title = `HPH_CP17_Dashboard_${new Date().toISOString().split('T')[0]}`;
    
    // Attendre un peu pour que les styles soient appliqués
    await new Promise(resolve => setTimeout(resolve, 300));
    
    // Lancer l'impression (qui permet de sauvegarder en PDF)
    window.print();
    
    // Restaurer après impression
    setTimeout(() => {
        // Restaurer les affichages
        elementsToHide.forEach((el, index) => {
            if (el) el.style.display = originalDisplayStates[index];
        });
        
        // Restaurer les backgrounds des KPI
        kpiCards.forEach((card, index) => {
            card.style.background = originalBackgrounds[index];
            card.style.removeProperty('background');
        });
        
        // Supprimer le titre temporaire
        const tempTitre = document.getElementById('pdf-titre-temp');
        if (tempTitre) tempTitre.remove();
        
        // Restaurer le titre
        document.title = originalTitle;
    }, 1000);
}

// Initialisation
document.addEventListener('DOMContentLoaded', function() {
    console.log('✅ Dashboard HPH CP17 initialisé');
    chargerObjectifs();
    chargerDonnees();
});
