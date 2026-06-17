// ============================================
// HPH CP17 - Dashboard Production
// Version Power Automate (fetch direct, sans authentification Microsoft)
// ============================================

// ============================================
// CONFIGURATION POWER AUTOMATE
// ============================================
// URL du flux de lecture HPH_CP17_Lecture_Donnees — renvoie directement
// {saisie: [...], iot: [...], empreintes: [...]} en JSON, sans authentification
// Microsoft requise (le flux est déclenché par une requête HTTP GET simple).
const URL_FLUX_LECTURE = "https://defaultefdc051fe1144edaa7c5efec56da13.b1.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/0c10446ba39c47b8a97cdf0d6ebf1d60/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=-7L4tPS29G5bK6TzgiCkY8-8k9xW6pJRxBXYZY_lzh8";

// Variables globales
let donneesCompletes = [];
let donneesFiltrees = [];
let donneesIoT = [];
let donneesEmpreintes = [];
let charts = {};

// Objectifs par défaut
let objectifs = {
    prodEquipe: 5000,
    tauxRebuts: 5,
    trs: 85
};

// ============================================
// CHARGEMENT DES DONNÉES
// ============================================
async function chargerDonnees() {
    const loadingMessage = document.getElementById('loadingMessage');
    const errorMessage = document.getElementById('errorMessage');

    loadingMessage.innerHTML = `<div class="spinner"></div><p>Chargement des données...</p>`;
    loadingMessage.classList.add('show');
    errorMessage.style.display = 'none';

    try {
        const response = await fetch(URL_FLUX_LECTURE);
        if (!response.ok) {
            throw new Error(`Power Automate erreur ${response.status}`);
        }
        const data = await response.json();

        donneesCompletes = data.saisie || [];
        donneesIoT = data.iot || [];
        donneesEmpreintes = data.empreintes || [];
        donneesFiltrees = [...donneesCompletes];

        console.log('✅ Données saisie:', donneesCompletes.length, 'lignes');
        console.log('✅ Données IoT:', donneesIoT.length, 'lignes');
        console.log('✅ Empreintes:', donneesEmpreintes.length, 'références');

        initialiserFiltresDates();
        appliquerFiltreEquipeParDefaut();
        mettreAJourStatutLive();

        // Rafraîchissement automatique toutes les 30 minutes (réduit la consommation
        // de requêtes Power Automate ; bouton "Actualiser" disponible pour un refresh manuel)
        setTimeout(chargerDonnees, 1800000);

    } catch (error) {
        console.error('❌ Erreur lors du chargement:', error);
        errorMessage.style.display = 'block';
        errorMessage.innerHTML = `<p>❌ Erreur lors du chargement des données : ${error.message}</p><button onclick="chargerDonnees()" class="btn btn-primary">🔄 Réessayer</button>`;
    } finally {
        loadingMessage.classList.remove('show');
    }
}

// ============================================
// HORAIRES PAR JOUR
// ============================================
function getHorairesEquipe(dateStr) {
    const date = new Date(dateStr);
    const jour = date.getDay();

    if (jour === 1) {
        return { Matin: [6, 13], Soir: [13, 21], Nuit: [21, 29] };
    } else if (jour >= 2 && jour <= 4) {
        return { Matin: [5, 13], Soir: [13, 21], Nuit: [21, 29] };
    } else if (jour === 5) {
        return { Matin: [5, 13], Soir: [13, 20], Nuit: [20, 28] };
    } else {
        return { Matin: [5, 13], Soir: [13, 21], Nuit: null };
    }
}

function getDureeEquipe(equipe, dateStr) {
    const horaires = getHorairesEquipe(dateStr);
    if (!horaires[equipe]) return 480;
    const [debut, fin] = horaires[equipe];
    return (fin - debut) * 60;
}

// ============================================
// CONVERSION DATE EXCEL
// ============================================
function convertirDateExcel(valeur) {
    if (typeof valeur === 'string' && (valeur.includes('/') || valeur.includes('-'))) {
        return valeur;
    }
    let valeurNumerique = valeur;
    if (typeof valeur === 'string' && !isNaN(valeur) && valeur.trim() !== '') {
        valeurNumerique = parseFloat(valeur);
    }
    if (typeof valeurNumerique === 'number' && valeurNumerique > 1000) {
        const dateExcelEpoch = new Date(1899, 11, 30);
        const dateMs = dateExcelEpoch.getTime() + (valeurNumerique * 86400000);
        const dateObj = new Date(dateMs);
        const annee = dateObj.getFullYear();
        const mois = String(dateObj.getMonth() + 1).padStart(2, '0');
        const jour = String(dateObj.getDate()).padStart(2, '0');
        return `${annee}-${mois}-${jour}`;
    }
    return valeur;
}

// ============================================
// GESTION EMPREINTES — saisie manuelle systématique (en attendant la licence
// Power Automate et la recherche automatique référence → nb empreintes)
// ============================================
function getNbEmpreintes(reference) {
    const champ = document.getElementById('nbEmpreintesManuel');
    const valeur = champ ? parseInt(champ.value) : NaN;
    return (!isNaN(valeur) && valeur > 0) ? valeur : 1;
}

// ============================================
// STATUT LIVE MACHINE (BANDEAU IoT)
// ============================================
function mettreAJourStatutLive() {
    if (!donneesIoT.length) {
        document.getElementById('statutLive').style.display = 'none';
        return;
    }

    const dernierEvenement = donneesIoT[donneesIoT.length - 1];
    const etat = (dernierEvenement['Etat'] || '').toString().trim().toUpperCase();

    const reference = document.getElementById('referenceObjectif') ?
        document.getElementById('referenceObjectif').value.trim().toUpperCase() : '';
    const nbEmpreintes = getNbEmpreintes(reference);

    let derniereCadenceBrute = null;
    for (let i = donneesIoT.length - 1; i >= 0; i--) {
        const ev = donneesIoT[i];
        if ((ev['Etat'] || '').toString().trim().toUpperCase() === 'CADENCE') {
            derniereCadenceBrute = parseFloat(ev['Cadence'] || 0);
            break;
        }
    }

    const cadencePieces = derniereCadenceBrute !== null ?
        (derniereCadenceBrute * nbEmpreintes).toFixed(1) : null;

    const heure = parseInt(dernierEvenement['Heure'] || 0);
    const minute = parseInt(dernierEvenement['Minute'] || 0);
    const maintenant = new Date();
    const minutesDepuis = (maintenant.getHours() * 60 + maintenant.getMinutes()) - (heure * 60 + minute);

    const statutDiv = document.getElementById('statutLive');
    statutDiv.style.display = 'flex';

    if (etat === 'ALLUMAGE' || etat === 'ON') {
        statutDiv.className = 'statut-live statut-marche';
        document.getElementById('statutIcon').textContent = '🟢';
        document.getElementById('statutTexte').textContent = 'EN MARCHE';
        document.getElementById('statutDepuis').textContent = minutesDepuis >= 0 ? `depuis ${minutesDepuis} min` : '';
        document.getElementById('statutCadence').textContent = cadencePieces ?
            `Cadence : ${cadencePieces} pcs/min` : '';
    } else if (etat === 'EXTINCTION' || etat === 'OFF') {
        statutDiv.className = 'statut-live statut-arret';
        document.getElementById('statutIcon').textContent = '🔴';
        document.getElementById('statutTexte').textContent = 'À L\'ARRÊT';
        document.getElementById('statutDepuis').textContent = minutesDepuis >= 0 ? `depuis ${minutesDepuis} min` : '';
        document.getElementById('statutCadence').textContent = cadencePieces ?
            `Dernière cadence : ${cadencePieces} pcs/min` : '';
    } else {
        statutDiv.style.display = 'none';
    }
}

// ============================================
// CALCUL CADENCE IoT PAR HEURE
// ============================================
function getCadenceIoTParHeure(dateStr, equipe) {
    if (!donneesIoT.length) return {};

    const cadenceParHeure = {};
    const horaires = getHorairesEquipe(dateStr);
    const [debutEquipe, finEquipe] = horaires[equipe] || [0, 24];

    const reference = document.getElementById('referenceObjectif') ?
        document.getElementById('referenceObjectif').value.trim().toUpperCase() : '';
    const nbEmpreintes = getNbEmpreintes(reference);

    const evenementsJour = donneesIoT.filter(ev => {
        const evDate = convertirDateExcel(ev['Date'] || '');
        let evDateISO = evDate;
        if (evDate.includes('/')) {
            const parts = evDate.split('/');
            evDateISO = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
        }
        return evDateISO === dateStr &&
               (ev['Etat'] || '').toString().trim().toUpperCase() === 'CADENCE';
    });

    evenementsJour.forEach(ev => {
        const heure = parseInt(ev['Heure'] || 0);
        const heureEquipe = heure - debutEquipe + 1;
        if (heureEquipe >= 1 && heureEquipe <= 8) {
            const label = `${heureEquipe}h`;
            if (!cadenceParHeure[label]) cadenceParHeure[label] = [];
            const cadenceBrute = parseFloat(ev['Cadence'] || 0);
            cadenceParHeure[label].push(cadenceBrute * nbEmpreintes);
        }
    });

    const resultat = {};
    Object.keys(cadenceParHeure).forEach(h => {
        const vals = cadenceParHeure[h];
        resultat[h] = vals.reduce((a, b) => a + b, 0) / vals.length;
    });
    return resultat;
}

// ============================================
// CADENCE DE RÉFÉRENCE RÉELLE (remplace l'ancien objectif fixe "cadence nominale")
//   Pour une ligne Saisie_Horaire donnée (date + équipe + heure), renvoie la
//   cadence IoT réelle mesurée par l'Opta (cycles/min × nb empreintes) pour
//   cette heure précise. Sert de référence pour le calcul du TRS/performance,
//   à la place d'une valeur théorique saisie manuellement.
// ============================================
function getCadenceReferenceLigne(ligne) {
    let dateStr = convertirDateExcel(ligne['Date'] || '');
    if (dateStr.includes('/')) {
        const parts = dateStr.split('/');
        dateStr = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
    }
    const equipe = ligne['Équipe'] || ligne['Equipe'] || 'Matin';
    const horaires = getHorairesEquipe(dateStr);
    const [debutEquipe] = horaires[equipe] || [0, 24];

    const heureAbs = parseInt(ligne['Heure'] || 0);
    const heureEquipe = heureAbs - debutEquipe + 1;
    const label = `${heureEquipe}h`;

    const cadenceParHeure = getCadenceIoTParHeure(dateStr, equipe);
    return cadenceParHeure[label] || null;  // null si pas de mesure IoT pour cette heure
}

// ============================================
// CALCUL TEMPS MARCHE IoT PAR HEURE
// ============================================
function getTempsMarcheIoTParHeure(dateStr, equipe) {
    if (!donneesIoT.length) return {};

    const tempsMarcheParHeure = {};
    const horaires = getHorairesEquipe(dateStr);
    const [debutEquipe] = horaires[equipe] || [0, 24];

    const evenementsJour = donneesIoT.filter(ev => {
        const evDate = convertirDateExcel(ev['Date'] || '');
        let evDateISO = evDate;
        if (evDate.includes('/')) {
            const parts = evDate.split('/');
            evDateISO = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
        }
        const etat = (ev['Etat'] || '').toString().trim().toUpperCase();
        return evDateISO === dateStr && (etat === 'ALLUMAGE' || etat === 'EXTINCTION' || etat === 'ON' || etat === 'OFF');
    }).sort((a, b) => {
        const tA = parseInt(a['Heure'] || 0) * 3600 + parseInt(a['Minute'] || 0) * 60 + parseInt(a['Seconde'] || 0);
        const tB = parseInt(b['Heure'] || 0) * 3600 + parseInt(b['Minute'] || 0) * 60 + parseInt(b['Seconde'] || 0);
        return tA - tB;
    });

    let dernierAllumage = null;
    evenementsJour.forEach(ev => {
        const etat = (ev['Etat'] || '').toString().trim().toUpperCase();
        const heureAbs = parseInt(ev['Heure'] || 0);
        const minuteAbs = parseInt(ev['Minute'] || 0);
        const secondeAbs = parseInt(ev['Seconde'] || 0);
        const tempsEnSecondes = heureAbs * 3600 + minuteAbs * 60 + secondeAbs;

        if (etat === 'ALLUMAGE' || etat === 'ON') {
            dernierAllumage = tempsEnSecondes;
        } else if ((etat === 'EXTINCTION' || etat === 'OFF') && dernierAllumage !== null) {
            const dureeMarche = tempsEnSecondes - dernierAllumage;
            const heureEquipeDebut = Math.floor(dernierAllumage / 3600) - debutEquipe + 1;
            const label = `${heureEquipeDebut}h`;
            if (!tempsMarcheParHeure[label]) tempsMarcheParHeure[label] = 0;
            tempsMarcheParHeure[label] += dureeMarche / 60;
            dernierAllumage = null;
        }
    });
    return tempsMarcheParHeure;
}

// ============================================
// GRAPHIQUE CADENCE IoT PAR HEURE
// ============================================
function afficherGraphiqueCadenceIoT() {
    const container = document.getElementById('chartCadenceIoT');
    if (!container) return;
    const ctx = container.getContext('2d');
    if (charts.cadenceIoT) charts.cadenceIoT.destroy();
    if (!donneesFiltrees.length) return;

    const premiereLigne = donneesFiltrees[0];
    let dateStr = convertirDateExcel(premiereLigne['Date'] || '');
    if (dateStr.includes('/')) {
        const parts = dateStr.split('/');
        dateStr = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
    }
    const equipe = premiereLigne['Équipe'] || premiereLigne['Equipe'] || 'Matin';
    const cadenceParHeure = getCadenceIoTParHeure(dateStr, equipe);
    const heures = ['1h','2h','3h','4h','5h','6h','7h','8h'];
    const cadenceData = heures.map(h => cadenceParHeure[h] || null);

    charts.cadenceIoT = new Chart(ctx, {
        type: 'line',
        data: {
            labels: heures,
            datasets: [
                {
                    label: 'Cadence réelle (pcs/min)',
                    data: cadenceData,
                    borderColor: 'rgba(54, 162, 235, 1)',
                    backgroundColor: 'rgba(54, 162, 235, 0.1)',
                    borderWidth: 3, pointRadius: 6,
                    pointBackgroundColor: 'rgba(54, 162, 235, 1)',
                    fill: true, tension: 0.3, spanGaps: true
                }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: true,
            scales: {
                y: { beginAtZero: true, title: { display: true, text: 'Pièces / minute' } },
                x: { title: { display: true, text: 'Heure d\'équipe' } }
            },
            plugins: {
                legend: { position: 'bottom' },
                tooltip: { callbacks: { label: ctx => ctx.parsed.y === null ? 'Pas de données' : `Cadence : ${ctx.parsed.y.toFixed(1)} pcs/min` } }
            }
        }
    });
}

// ============================================
// GRAPHIQUE RÉPARTITION MARCHE / ARRÊT PAR HEURE (colonnes empilées)
// ============================================
function afficherGraphiqueMarcheArretHeure() {
    const container = document.getElementById('chartMarcheArretHeure');
    if (!container) return;
    const ctx = container.getContext('2d');
    if (charts.marcheArretHeure) charts.marcheArretHeure.destroy();
    if (!donneesFiltrees.length) return;

    const premiereLigne = donneesFiltrees[0];
    let dateStr = convertirDateExcel(premiereLigne['Date'] || '');
    if (dateStr.includes('/')) {
        const parts = dateStr.split('/');
        dateStr = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
    }
    const equipe = premiereLigne['Équipe'] || premiereLigne['Equipe'] || 'Matin';
    const tempsMarcheParHeure = getTempsMarcheIoTParHeure(dateStr, equipe);

    // Durées déclaratives d'arrêt par heure d'équipe (fallback si pas de données IoT pour l'heure)
    // ligne['Heure'] est l'heure d'équipe relative (1, 2, 3...) comme utilisé dans les autres graphiques horaires
    const arretsDeclaratifsParHeure = {};
    donneesFiltrees.forEach(ligne => {
        const heureLabel = `${ligne['Heure'] || '?'}h`;
        const arrets = parseInt(ligne['Équipement Durée'] || ligne['Équipement_x0020_Durée'] || 0)
            + parseInt(ligne['Qualité Durée'] || ligne['Qualité_x0020_Durée'] || 0)
            + parseInt(ligne['Organisation Durée'] || ligne['Organisation_x0020_Durée'] || 0)
            + parseInt(ligne['Autres Durée'] || ligne['Autres_x0020_Durée'] || 0);
        arretsDeclaratifsParHeure[heureLabel] = (arretsDeclaratifsParHeure[heureLabel] || 0) + arrets;
    });

    const heures = ['1h','2h','3h','4h','5h','6h','7h','8h'];
    const dureeEquipeMin = getDureeEquipe(equipe, dateStr) / 8;  // durée moyenne d'une heure d'équipe (gère le cas vendredi 4h)

    const margeMin = [];
    const margeArret = [];
    heures.forEach(h => {
        const tempsMarcheIoT = tempsMarcheParHeure[h];
        if (tempsMarcheIoT !== undefined) {
            // Source IoT disponible pour cette heure
            const marche = Math.min(tempsMarcheIoT, dureeEquipeMin);
            margeMin.push(Math.round(marche * 10) / 10);
            margeArret.push(Math.round(Math.max(0, dureeEquipeMin - marche) * 10) / 10);
        } else if (arretsDeclaratifsParHeure[h] !== undefined) {
            // Repli sur les durées déclaratives (Saisie_Horaire) si pas de données IoT
            const arret = arretsDeclaratifsParHeure[h];
            margeArret.push(Math.round(Math.min(arret, dureeEquipeMin) * 10) / 10);
            margeMin.push(Math.round(Math.max(0, dureeEquipeMin - arret) * 10) / 10);
        } else {
            // Aucune donnée du tout pour cette heure
            margeMin.push(null);
            margeArret.push(null);
        }
    });

    charts.marcheArretHeure = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: heures,
            datasets: [
                {
                    label: 'Marche (min)',
                    data: margeMin,
                    backgroundColor: 'rgba(75, 192, 75, 0.7)',
                    borderColor: 'rgba(75, 192, 75, 1)',
                    borderWidth: 1,
                    stack: 'stack1'
                },
                {
                    label: 'Arrêt (min)',
                    data: margeArret,
                    backgroundColor: 'rgba(255, 99, 99, 0.7)',
                    borderColor: 'rgba(255, 99, 99, 1)',
                    borderWidth: 1,
                    stack: 'stack1'
                }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: true,
            scales: {
                y: { beginAtZero: true, max: Math.ceil(dureeEquipeMin), stacked: true, title: { display: true, text: 'Minutes' } },
                x: { stacked: true, title: { display: true, text: 'Heure d\'équipe' } }
            },
            plugins: {
                legend: { position: 'bottom' },
                tooltip: { callbacks: { label: ctx => ctx.parsed.y === null ? 'Pas de données' : `${ctx.dataset.label} : ${ctx.parsed.y} min` } }
            }
        }
    });
}
function calculerPerformanceIoT(heure, donneeHeure) {
    if (!donneesIoT.length) return null;
    const premiereLigne = donneesFiltrees[0];
    let dateStr = convertirDateExcel(premiereLigne['Date'] || '');
    if (dateStr.includes('/')) {
        const parts = dateStr.split('/');
        dateStr = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`;
    }
    const equipe = premiereLigne['Équipe'] || premiereLigne['Equipe'] || 'Matin';
    const tempsMarcheParHeure = getTempsMarcheIoTParHeure(dateStr, equipe);
    const tempsMarche = tempsMarcheParHeure[heure] || null;
    if (tempsMarche === null) return null;

    // Cadence de référence réelle (IoT) pour cette heure précise — plus de valeur théorique fixe
    const cadenceParHeure = getCadenceIoTParHeure(dateStr, equipe);
    const cadenceReelle = cadenceParHeure[heure] || null;
    if (cadenceReelle === null) return null;

    const prodTotale = donneeHeure.bonne + donneeHeure.rebuts;
    const prodTheorique = cadenceReelle * tempsMarche;
    if (prodTheorique <= 0) return null;
    return Math.min((prodTotale / prodTheorique) * 100, 150);
}

// ============================================
// AFFICHAGE DASHBOARD
// ============================================
function afficherDashboard() {
    if (donneesFiltrees.length === 0) {
        document.getElementById('errorMessage').style.display = 'block';
        document.getElementById('errorMessage').innerHTML = '<p>ℹ️ Aucune donnée disponible.</p>';
        return;
    }
    calculerKPIs();
    afficherGraphiques();
    afficherTableau();
    mettreAJourStatutLive();
}

// ============================================
// KPIs
// ============================================
function calculerKPIs() {
    let totalProdBonne = 0, totalRebuts = 0, totalTempsArret = 0;
    let prodMaxGlobale = 0;
    let nbLignesAvecCadence = 0;

    donneesFiltrees.forEach(ligne => {
        totalProdBonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
        totalRebuts += parseInt(ligne['Rebuts'] || 0);
        totalTempsArret += parseInt(ligne['Équipement Durée'] || ligne['Équipement_x0020_Durée'] || 0)
            + parseInt(ligne['Qualité Durée'] || ligne['Qualité_x0020_Durée'] || 0)
            + parseInt(ligne['Organisation Durée'] || ligne['Organisation_x0020_Durée'] || 0)
            + parseInt(ligne['Autres Durée'] || ligne['Autres_x0020_Durée'] || 0);

        // Cadence de référence réelle (IoT) propre à cette ligne (date+équipe+heure)
        const cadenceRef = getCadenceReferenceLigne(ligne);
        if (cadenceRef !== null) {
            prodMaxGlobale += cadenceRef * 60;  // production théorique sur 1h à cette cadence
            nbLignesAvecCadence++;
        }
    });

    const totalProduction = totalProdBonne + totalRebuts;
    const tauxRebuts = totalProduction > 0 ? ((totalRebuts / totalProduction) * 100).toFixed(2) : 0;
    const trsGlobal = prodMaxGlobale > 0 ? ((totalProduction / prodMaxGlobale) * 100).toFixed(1) : 0;

    document.getElementById('kpiProdTotal').textContent = totalProduction.toLocaleString();
    document.getElementById('kpiProdBonne').textContent = totalProdBonne.toLocaleString();
    document.getElementById('kpiRebuts').textContent = totalRebuts.toLocaleString();
    document.getElementById('kpiTauxRebuts').textContent = tauxRebuts + '%';
    document.getElementById('kpiTempsArret').textContent = totalTempsArret + ' min';
    document.getElementById('kpiTRSGlobal').textContent = nbLignesAvecCadence > 0 ? trsGlobal + '%' : 'N/A';

    compareAvecObjectifs(totalProduction, tauxRebuts, totalTempsArret, parseFloat(trsGlobal) || 0);
}

function compareAvecObjectifs(totalProduction, tauxRebuts, totalTempsArret, trsGlobal) {
    const tauxRebutsNum = parseFloat(tauxRebuts);
    const equipesUniques = new Set();
    donneesFiltrees.forEach(ligne => { if (ligne['Équipe']) equipesUniques.add(ligne['Équipe']); });
    const nombreEquipes = equipesUniques.size;
    const objectifTotal = objectifs.prodEquipe * nombreEquipes;

    const kpiCardProd = document.getElementById('kpiCardProdTotal');
    const kpiObjectifProd = document.getElementById('kpiObjectifProd');
    const kpiStatusProd = document.getElementById('kpiStatusProd');

    kpiObjectifProd.textContent = nombreEquipes > 1 ?
        `Objectif: ${objectifTotal.toLocaleString()} (${nombreEquipes} équipes × ${objectifs.prodEquipe.toLocaleString()})` :
        `Objectif: ${objectifTotal.toLocaleString()} (${nombreEquipes} équipe)`;

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

    const kpiCardRebuts = document.getElementById('kpiCardTauxRebuts');
    document.getElementById('kpiObjectifRebuts').textContent = `Objectif max: ${objectifs.tauxRebuts}%`;
    if (tauxRebutsNum <= objectifs.tauxRebuts) {
        document.getElementById('kpiStatusRebuts').textContent = '✅ Objectif respecté';
        document.getElementById('kpiStatusRebuts').className = 'kpi-status atteint';
        kpiCardRebuts.className = 'kpi-card success';
    } else {
        document.getElementById('kpiStatusRebuts').textContent = '❌ Objectif dépassé';
        document.getElementById('kpiStatusRebuts').className = 'kpi-status non-atteint';
        kpiCardRebuts.className = 'kpi-card danger';
    }

    const kpiCardTRSGlobal = document.getElementById('kpiCardTRSGlobal');
    document.getElementById('kpiObjectifTRSGlobal').textContent = `Objectif: ${objectifs.trs}%`;
    if (trsGlobal >= objectifs.trs) {
        document.getElementById('kpiStatusTRSGlobal').textContent = '✅ Objectif atteint';
        document.getElementById('kpiStatusTRSGlobal').className = 'kpi-status atteint';
        kpiCardTRSGlobal.className = 'kpi-card success';
    } else {
        document.getElementById('kpiStatusTRSGlobal').textContent = `⚠️ ${trsGlobal}% (objectif: ${objectifs.trs}%)`;
        document.getElementById('kpiStatusTRSGlobal').className = 'kpi-status non-atteint';
        kpiCardTRSGlobal.className = 'kpi-card warning';
    }
}

// ============================================
// GRAPHIQUES
// ============================================
function afficherGraphiques() {
    const datesUniques = new Set();
    const equipesUniques = new Set();
    donneesFiltrees.forEach(ligne => {
        let date = convertirDateExcel(ligne['Date'] || '');
        if (date.includes('/')) { const parts = date.split('/'); date = `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`; }
        datesUniques.add(date);
        const equipe = ligne['Équipe'] || ligne['Equipe'];
        if (equipe) equipesUniques.add(equipe);
    });

    if (datesUniques.size === 1 && equipesUniques.size === 1) {
        document.getElementById('graphiques-horaires').style.display = 'block';
        document.getElementById('graphiques-journaliers').style.display = 'none';
        afficherGraphiqueProductionHeure();
        afficherGraphiqueTRS();
        afficherGraphiquePerformance();
        afficherGraphiqueProductionEquipe();
        afficherGraphiqueCadenceIoT();
        afficherGraphiqueMarcheArretHeure();
        document.getElementById('graphiques-horaires').style.display = 'none';
        document.getElementById('graphiques-journaliers').style.display = 'block';
        afficherGraphiqueProductionJour();
        afficherGraphiqueTRSJour();
        afficherGraphiquePerformanceJour();
        afficherGraphiqueTempsNonJustifieJour();
    }
    afficherGraphiqueDureeArrets();
}

function afficherGraphiqueProductionHeure() {
    const ctx = document.getElementById('chartProductionHeure').getContext('2d');
    if (charts.productionHeure) charts.productionHeure.destroy();
    const dataParHeure = {};
    donneesFiltrees.forEach(ligne => {
        const heure = ligne['Heure'] || '?';
        if (!dataParHeure[heure]) dataParHeure[heure] = { bonne: 0 };
        dataParHeure[heure].bonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
    });
    const heures = Object.keys(dataParHeure).sort((a, b) => parseInt(a) - parseInt(b));
    const prodBonne = heures.map(h => dataParHeure[h].bonne);
    const objectifParHeure = objectifs.prodEquipe / 8;
    const backgroundColor = prodBonne.map(val => val >= objectifParHeure ? 'rgba(75, 192, 75, 0.6)' : 'rgba(255, 99, 99, 0.6)');
    const borderColor = prodBonne.map(val => val >= objectifParHeure ? 'rgba(75, 192, 75, 1)' : 'rgba(255, 99, 99, 1)');
    charts.productionHeure = new Chart(ctx, {
        type: 'bar',
        data: { labels: heures, datasets: [
            { label: 'Production Bonne', data: prodBonne, backgroundColor, borderColor, borderWidth: 2, order: 2 },
            { label: 'Objectif/Heure', data: heures.map(() => objectifParHeure), type: 'line', borderColor: 'rgba(255, 193, 7, 1)', borderWidth: 3, borderDash: [10, 5], pointRadius: 5, fill: false, order: 1, tension: 0 }
        ]},
        options: { responsive: true, maintainAspectRatio: true, scales: { y: { beginAtZero: true }, x: {} }, plugins: { legend: { position: 'bottom' } } }
    });
}

function afficherGraphiqueTRS() {
    const ctx = document.getElementById('chartTRS').getContext('2d');
    if (charts.trs) charts.trs.destroy();
    const dataParHeure = {};
    donneesFiltrees.forEach(ligne => {
        const heure = ligne['Heure'] || '?';
        if (!dataParHeure[heure]) dataParHeure[heure] = { bonne: 0, rebuts: 0, cadenceRef: getCadenceReferenceLigne(ligne) };
        dataParHeure[heure].bonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
        dataParHeure[heure].rebuts += parseInt(ligne['Rebuts'] || 0);
    });
    const heures = Object.keys(dataParHeure).sort((a, b) => parseInt(a) - parseInt(b));
    const trsData = heures.map(h => {
        const d = dataParHeure[h];
        const prodTotale = d.bonne + d.rebuts;
        if (!d.cadenceRef) return 0;  // pas de mesure IoT pour cette heure → pas de TRS calculable
        return Math.min(prodTotale > 0 ? (prodTotale / (d.cadenceRef * 60)) * 100 : 0, 150);
    });
    const backgroundColor = trsData.map(val => val >= objectifs.trs ? 'rgba(75, 192, 192, 0.6)' : 'rgba(255, 99, 132, 0.6)');
    charts.trs = new Chart(ctx, {
        type: 'bar',
        data: { labels: heures, datasets: [
            { label: 'TRS Réalisé (%)', data: trsData, backgroundColor, borderColor: backgroundColor.map(c => c.replace('0.6','1')), borderWidth: 2, order: 2 },
            { label: 'Objectif TRS (%)', data: heures.map(() => objectifs.trs), type: 'line', borderColor: 'rgba(255, 193, 7, 1)', borderWidth: 3, borderDash: [10, 5], pointRadius: 5, fill: false, order: 1, tension: 0 }
        ]},
        options: { responsive: true, maintainAspectRatio: true, scales: { y: { beginAtZero: true, max: 170, ticks: { callback: v => v + '%' } } }, plugins: { legend: { position: 'bottom' } } }
    });
}

function afficherGraphiquePerformance() {
    const ctx = document.getElementById('chartPerformance').getContext('2d');
    if (charts.performance) charts.performance.destroy();
    const dataParHeure = {};
    donneesFiltrees.forEach(ligne => {
        const heure = ligne['Heure'] || '?';
        if (!dataParHeure[heure]) dataParHeure[heure] = { bonne: 0, rebuts: 0, arrets: 0, cadenceRef: getCadenceReferenceLigne(ligne) };
        dataParHeure[heure].bonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
        dataParHeure[heure].rebuts += parseInt(ligne['Rebuts'] || 0);
        dataParHeure[heure].arrets += parseInt(ligne['Équipement Durée'] || 0) + parseInt(ligne['Qualité Durée'] || 0) + parseInt(ligne['Organisation Durée'] || 0) + parseInt(ligne['Autres Durée'] || 0);
    });
    const heures = Object.keys(dataParHeure).sort((a, b) => parseInt(a) - parseInt(b));
    const performanceData = heures.map(h => {
        const perfIoT = calculerPerformanceIoT(h, dataParHeure[h]);
        if (perfIoT !== null) return perfIoT;
        const d = dataParHeure[h];
        if (!d.cadenceRef) return 0;  // pas de cadence IoT disponible pour cette heure
        const prodTotale = d.bonne + d.rebuts;
        const tempsProductif = 60 - d.arrets;
        if (tempsProductif <= 0) return 0;
        return Math.min((prodTotale / (d.cadenceRef * tempsProductif)) * 100, 150);
    });
    const sourceIoT = donneesIoT.length > 0;
    const backgroundColor = performanceData.map(val => val >= 100 ? 'rgba(75, 192, 75, 0.6)' : val >= 85 ? 'rgba(255, 193, 7, 0.6)' : 'rgba(255, 99, 132, 0.6)');
    charts.performance = new Chart(ctx, {
        type: 'bar',
        data: { labels: heures, datasets: [
            { label: `Performance ${sourceIoT ? '(IoT 🔵)' : '(Déclaratif)'}`, data: performanceData, backgroundColor, borderColor: backgroundColor.map(c => c.replace('0.6','1')), borderWidth: 2, order: 2 },
            { label: 'Objectif (100%)', data: heures.map(() => 100), type: 'line', borderColor: 'rgba(255, 193, 7, 1)', borderWidth: 3, borderDash: [10, 5], pointRadius: 5, fill: false, order: 1, tension: 0 }
        ]},
        options: { responsive: true, maintainAspectRatio: true, scales: { y: { beginAtZero: true, max: 170, ticks: { callback: v => v + '%' } } }, plugins: { legend: { position: 'bottom' } } }
    });
}

function afficherGraphiqueProductionEquipe() {
    const ctx = document.getElementById('chartProductionEquipe').getContext('2d');
    if (charts.productionEquipe) charts.productionEquipe.destroy();
    const dataParHeure = {};
    donneesFiltrees.forEach(ligne => {
        const heure = ligne['Heure'] || '?';
        if (!dataParHeure[heure]) dataParHeure[heure] = { prodBonne: 0, rebuts: 0, arrets: 0, cadenceRef: getCadenceReferenceLigne(ligne) };
        dataParHeure[heure].prodBonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
        dataParHeure[heure].rebuts += parseInt(ligne['Rebuts'] || 0);
        dataParHeure[heure].arrets += parseInt(ligne['Équipement Durée'] || 0) + parseInt(ligne['Qualité Durée'] || 0) + parseInt(ligne['Organisation Durée'] || 0) + parseInt(ligne['Autres Durée'] || 0);
    });
    const heures = Object.keys(dataParHeure).sort((a, b) => parseInt(a) - parseInt(b));
    const tempsNonJustifieData = heures.map(h => {
        const data = dataParHeure[h];
        if (!data.cadenceRef) return 0;  // pas de cadence IoT disponible pour cette heure
        const prodTotale = data.prodBonne + data.rebuts;
        const prodAttendue = data.cadenceRef * 60;
        if (prodTotale >= prodAttendue) return 0;
        return Math.max(0, ((prodAttendue - prodTotale) / data.cadenceRef) - data.arrets);
    });
    const backgroundColor = tempsNonJustifieData.map(val => val === 0 ? 'rgba(75, 192, 75, 0.6)' : val > 5 ? 'rgba(255, 99, 99, 0.6)' : 'rgba(255, 193, 7, 0.6)');
    charts.productionEquipe = new Chart(ctx, {
        type: 'bar',
        data: { labels: heures, datasets: [{ label: 'Temps Non Justifié (min)', data: tempsNonJustifieData, backgroundColor, borderColor: backgroundColor.map(c => c.replace('0.6','1')), borderWidth: 2 }]},
        options: { responsive: true, maintainAspectRatio: true, scales: { y: { beginAtZero: true, title: { display: true, text: 'Minutes' } } }, plugins: { legend: { position: 'bottom' } } }
    });
}

function afficherGraphiqueDureeArrets() {
    const ctx = document.getElementById('chartDureeArrets').getContext('2d');
    if (charts.dureeArrets) charts.dureeArrets.destroy();
    let dureeEquipement = 0, dureeQualite = 0, dureeOrganisation = 0, dureeAutres = 0;
    donneesFiltrees.forEach(ligne => {
        dureeEquipement += parseInt(ligne['Équipement Durée'] || ligne['Équipement_x0020_Durée'] || 0);
        dureeQualite += parseInt(ligne['Qualité Durée'] || ligne['Qualité_x0020_Durée'] || 0);
        dureeOrganisation += parseInt(ligne['Organisation Durée'] || ligne['Organisation_x0020_Durée'] || 0);
        dureeAutres += parseInt(ligne['Autres Durée'] || ligne['Autres_x0020_Durée'] || 0);
    });
    charts.dureeArrets = new Chart(ctx, {
        type: 'bar',
        data: { labels: ['Équipement', 'Qualité', 'Organisation', 'Autres'], datasets: [{
            label: 'Durée (minutes)',
            data: [dureeEquipement, dureeQualite, dureeOrganisation, dureeAutres],
            backgroundColor: ['rgba(255, 159, 64, 0.6)', 'rgba(255, 99, 132, 0.6)', 'rgba(54, 162, 235, 0.6)', 'rgba(153, 102, 255, 0.6)'],
            borderColor: ['rgba(255, 159, 64, 1)', 'rgba(255, 99, 132, 1)', 'rgba(54, 162, 235, 1)', 'rgba(153, 102, 255, 1)'],
            borderWidth: 2
        }]},
        options: { responsive: true, maintainAspectRatio: true, scales: { y: { beginAtZero: true } }, plugins: { legend: { position: 'bottom' } } }
    });
}

function afficherGraphiqueProductionJour() {
    const ctx = document.getElementById('chartProductionJour').getContext('2d');
    if (charts.productionJour) charts.productionJour.destroy();
    const dataParJour = {};
    donneesFiltrees.forEach(ligne => {
        let date = convertirDateExcel(ligne['Date'] || '');
        if (date.includes('/')) { const p = date.split('/'); date = `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`; }
        if (!dataParJour[date]) dataParJour[date] = { bonne: 0 };
        dataParJour[date].bonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
    });
    const jours = Object.keys(dataParJour).sort();
    const prodBonneData = jours.map(j => dataParJour[j].bonne);
    const objectifJour = objectifs.prodEquipe;
    const backgroundColor = prodBonneData.map(val => val >= objectifJour ? 'rgba(75, 192, 75, 0.6)' : 'rgba(255, 99, 99, 0.6)');
    charts.productionJour = new Chart(ctx, {
        type: 'bar',
        data: { labels: jours.map(d => new Date(d).toLocaleDateString('fr-FR')), datasets: [
            { label: 'Production Bonne', data: prodBonneData, backgroundColor, borderColor: backgroundColor.map(c => c.replace('0.6','1')), borderWidth: 2, order: 2 },
            { label: 'Objectif/Jour', data: jours.map(() => objectifJour), type: 'line', borderColor: 'rgba(255, 193, 7, 1)', borderWidth: 3, borderDash: [10, 5], pointRadius: 5, fill: false, order: 1, tension: 0 }
        ]},
        options: { responsive: true, maintainAspectRatio: true, scales: { y: { beginAtZero: true } }, plugins: { legend: { position: 'bottom' } } }
    });
}

function afficherGraphiqueTRSJour() {
    const ctx = document.getElementById('chartTRSJour').getContext('2d');
    if (charts.trsJour) charts.trsJour.destroy();
    const dataParJour = {};
    donneesFiltrees.forEach(ligne => {
        let date = convertirDateExcel(ligne['Date'] || '');
        if (date.includes('/')) { const p = date.split('/'); date = `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`; }
        if (!dataParJour[date]) dataParJour[date] = { bonne: 0, rebuts: 0, prodMaxJour: 0 };
        dataParJour[date].bonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
        dataParJour[date].rebuts += parseInt(ligne['Rebuts'] || 0);
        const cadenceRef = getCadenceReferenceLigne(ligne);
        if (cadenceRef) dataParJour[date].prodMaxJour += cadenceRef * 60;
    });
    const jours = Object.keys(dataParJour).sort();
    const trsData = jours.map(j => {
        const d = dataParJour[j];
        const prodTotale = d.bonne + d.rebuts;
        if (d.prodMaxJour <= 0) return 0;  // pas de cadence IoT disponible ce jour-là
        return Math.min(prodTotale > 0 ? (prodTotale / d.prodMaxJour) * 100 : 0, 150);
    });
    const backgroundColor = trsData.map(val => val >= objectifs.trs ? 'rgba(75, 192, 192, 0.6)' : 'rgba(255, 99, 132, 0.6)');
    charts.trsJour = new Chart(ctx, {
        type: 'bar',
        data: { labels: jours.map(d => new Date(d).toLocaleDateString('fr-FR')), datasets: [
            { label: 'TRS Réalisé (%)', data: trsData, backgroundColor, borderColor: backgroundColor.map(c => c.replace('0.6','1')), borderWidth: 2, order: 2 },
            { label: 'Objectif TRS (%)', data: jours.map(() => objectifs.trs), type: 'line', borderColor: 'rgba(255, 193, 7, 1)', borderWidth: 3, borderDash: [10, 5], pointRadius: 5, fill: false, order: 1, tension: 0 }
        ]},
        options: { responsive: true, maintainAspectRatio: true, scales: { y: { beginAtZero: true, max: 170, ticks: { callback: v => v + '%' } } }, plugins: { legend: { position: 'bottom' } } }
    });
}

function afficherGraphiquePerformanceJour() {
    const ctx = document.getElementById('chartPerformanceJour').getContext('2d');
    if (charts.performanceJour) charts.performanceJour.destroy();
    const dataParJour = {};
    donneesFiltrees.forEach(ligne => {
        let date = convertirDateExcel(ligne['Date'] || '');
        if (date.includes('/')) { const p = date.split('/'); date = `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`; }
        if (!dataParJour[date]) dataParJour[date] = { bonne: 0, rebuts: 0, arrets: 0, prodMaxJour: 0 };
        dataParJour[date].bonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
        dataParJour[date].rebuts += parseInt(ligne['Rebuts'] || 0);
        dataParJour[date].arrets += parseInt(ligne['Équipement Durée'] || 0) + parseInt(ligne['Qualité Durée'] || 0) + parseInt(ligne['Organisation Durée'] || 0) + parseInt(ligne['Autres Durée'] || 0);
        const cadenceRef = getCadenceReferenceLigne(ligne);
        if (cadenceRef) dataParJour[date].prodMaxJour += cadenceRef * 60;
    });
    const jours = Object.keys(dataParJour).sort();
    const performanceData = jours.map(j => {
        const d = dataParJour[j];
        const prodTotale = d.bonne + d.rebuts;
        const tempsProductif = (8 * 60) - d.arrets;
        if (tempsProductif <= 0 || d.prodMaxJour <= 0) return 0;
        // prodMaxJour est déjà calculée sur les heures réelles à cadence IoT ; on la ramène
        // au temps productif réel (hors arrêts) pour rester cohérent avec l'ancienne formule.
        const prodTheoriqueProductif = d.prodMaxJour * (tempsProductif / (8 * 60));
        if (prodTheoriqueProductif <= 0) return 0;
        return Math.min((prodTotale / prodTheoriqueProductif) * 100, 150);
    });
    const backgroundColor = performanceData.map(val => val >= 100 ? 'rgba(75, 192, 75, 0.6)' : val >= 85 ? 'rgba(255, 193, 7, 0.6)' : 'rgba(255, 99, 132, 0.6)');
    charts.performanceJour = new Chart(ctx, {
        type: 'bar',
        data: { labels: jours.map(d => new Date(d).toLocaleDateString('fr-FR')), datasets: [
            { label: 'Performance (%)', data: performanceData, backgroundColor, borderColor: backgroundColor.map(c => c.replace('0.6','1')), borderWidth: 2, order: 2 },
            { label: 'Objectif (100%)', data: jours.map(() => 100), type: 'line', borderColor: 'rgba(255, 193, 7, 1)', borderWidth: 3, borderDash: [10, 5], pointRadius: 5, fill: false, order: 1, tension: 0 }
        ]},
        options: { responsive: true, maintainAspectRatio: true, scales: { y: { beginAtZero: true, max: 170, ticks: { callback: v => v + '%' } } }, plugins: { legend: { position: 'bottom' } } }
    });
}

function afficherGraphiqueTempsNonJustifieJour() {
    const ctx = document.getElementById('chartTempsNonJustifieJour').getContext('2d');
    if (charts.tempsNonJustifieJour) charts.tempsNonJustifieJour.destroy();
    const dataParJour = {};
    donneesFiltrees.forEach(ligne => {
        let date = convertirDateExcel(ligne['Date'] || '');
        if (date.includes('/')) { const p = date.split('/'); date = `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`; }
        if (!dataParJour[date]) dataParJour[date] = { prodBonne: 0, rebuts: 0, arrets: 0, prodMaxJour: 0, cadenceMoyJour: [] };
        dataParJour[date].prodBonne += parseInt(ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0);
        dataParJour[date].rebuts += parseInt(ligne['Rebuts'] || 0);
        dataParJour[date].arrets += parseInt(ligne['Équipement Durée'] || 0) + parseInt(ligne['Qualité Durée'] || 0) + parseInt(ligne['Organisation Durée'] || 0) + parseInt(ligne['Autres Durée'] || 0);
        const cadenceRef = getCadenceReferenceLigne(ligne);
        if (cadenceRef) { dataParJour[date].prodMaxJour += cadenceRef * 60; dataParJour[date].cadenceMoyJour.push(cadenceRef); }
    });
    const jours = Object.keys(dataParJour).sort();
    const tempsNonJustifieData = jours.map(j => {
        const d = dataParJour[j];
        if (d.prodMaxJour <= 0) return 0;  // pas de cadence IoT disponible ce jour-là
        const cadenceMoyenne = d.cadenceMoyJour.reduce((a,b) => a+b, 0) / d.cadenceMoyJour.length;
        const prodTotale = d.prodBonne + d.rebuts;
        if (prodTotale >= d.prodMaxJour) return 0;
        return Math.max(0, ((d.prodMaxJour - prodTotale) / cadenceMoyenne) - d.arrets);
    });
    const backgroundColor = tempsNonJustifieData.map(val => val === 0 ? 'rgba(75, 192, 75, 0.6)' : val <= 30 ? 'rgba(255, 193, 7, 0.6)' : 'rgba(255, 99, 132, 0.6)');
    charts.tempsNonJustifieJour = new Chart(ctx, {
        type: 'bar',
        data: { labels: jours.map(d => new Date(d).toLocaleDateString('fr-FR')), datasets: [{ label: 'Temps Non Justifié (min)', data: tempsNonJustifieData, backgroundColor, borderColor: backgroundColor.map(c => c.replace('0.6','1')), borderWidth: 2 }]},
        options: { responsive: true, maintainAspectRatio: true, scales: { y: { beginAtZero: true } }, plugins: { legend: { position: 'bottom' } } }
    });
}

// ============================================
// TABLEAU DÉTAILLÉ
// ============================================
function afficherTableau() {
    const tbody = document.getElementById('dataTableBody');
    tbody.innerHTML = '';
    donneesFiltrees.forEach(ligne => {
        const row = tbody.insertRow();
        let date = convertirDateExcel(ligne['Date'] || '');
        if (date) {
            try {
                const dateObj = new Date(date);
                if (!isNaN(dateObj.getTime())) {
                    date = `${String(dateObj.getDate()).padStart(2,'0')}/${String(dateObj.getMonth()+1).padStart(2,'0')}/${dateObj.getFullYear()}`;
                }
            } catch(e) {}
        }
        row.innerHTML = `
            <td>${date}</td>
            <td>${ligne['Référence'] || ligne['Reference'] || ''}</td>
            <td>${ligne['Jour'] || ''}</td>
            <td>${ligne['Équipe'] || ligne['Equipe'] || ''}</td>
            <td>${ligne['Heure'] || ''}</td>
            <td>${ligne['Prod Bonne'] || ligne['Prod_x0020_Bonne'] || 0}</td>
            <td>${ligne['Rebuts'] || 0}</td>
            <td>${ligne['Équipement'] || ''}</td>
            <td>${ligne['Équipement Durée'] || ligne['Équipement_x0020_Durée'] || 0}</td>
            <td>${ligne['Qualité'] || ''}</td>
            <td>${ligne['Qualité Durée'] || ligne['Qualité_x0020_Durée'] || 0}</td>
            <td>${ligne['Organisation'] || ''}</td>
            <td>${ligne['Organisation Durée'] || ligne['Organisation_x0020_Durée'] || 0}</td>
            <td>${ligne['Autres'] || ''}</td>
            <td>${ligne['Autres Durée'] || ligne['Autres_x0020_Durée'] || 0}</td>
            <td>${ligne['Commentaire'] || ''}</td>
        `;
    });
}

// ============================================
// FILTRES
// ============================================
function initialiserFiltresDates() {
    const dateAujourdhui = new Date().toISOString().split('T')[0];
    document.getElementById('filterDateDebut').value = dateAujourdhui;
    document.getElementById('filterDateFin').value = dateAujourdhui;
}

function determinerEquipeEnCours() {
    const heure = new Date().getHours();
    const jour = new Date().getDay();
    let debutMatin = 5, debutSoir = 13, debutNuit = 21;
    if (jour === 1) { debutMatin = 6; }
    else if (jour === 5) { debutNuit = 20; }
    if (heure >= debutMatin && heure < debutSoir) return 'Matin';
    if (heure >= debutSoir && heure < debutNuit) return 'Soir';
    return 'Nuit';
}

function appliquerFiltreEquipeParDefaut() {
    const equipeEnCours = determinerEquipeEnCours();
    const dateAujourdhui = new Date().toISOString().split('T')[0];
    document.getElementById('filterEquipe').value = equipeEnCours;

    donneesFiltrees = donneesCompletes.filter(ligne => {
        const equipe = ligne['Équipe'];
        let ligneDate = convertirDateExcel(ligne['Date'] || '');
        if (ligneDate.includes('/')) { const p = ligneDate.split('/'); ligneDate = `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`; }
        return equipe === equipeEnCours && ligneDate === dateAujourdhui;
    });

    if (donneesFiltrees.length === 0) {
        const donneesEquipe = donneesCompletes.filter(l => l['Équipe'] === equipeEnCours);
        if (donneesEquipe.length > 0) {
            const dates = new Set(donneesEquipe.map(l => {
                let d = convertirDateExcel(l['Date'] || '');
                if (d.includes('/')) { const p = d.split('/'); d = `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`; }
                return d;
            }));
            const derniereDate = Array.from(dates).sort().reverse()[0];
            donneesFiltrees = donneesEquipe.filter(l => {
                let d = convertirDateExcel(l['Date'] || '');
                if (d.includes('/')) { const p = d.split('/'); d = `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`; }
                return d === derniereDate;
            });
        } else {
            donneesFiltrees = [...donneesCompletes];
        }
    }
    afficherDashboard();
}

function appliquerFiltres() {
    const filterDateDebut = document.getElementById('filterDateDebut').value;
    const filterDateFin = document.getElementById('filterDateFin').value;
    const filterEquipe = document.getElementById('filterEquipe').value;

    donneesFiltrees = donneesCompletes.filter(ligne => {
        let match = true;
        if (filterDateDebut || filterDateFin) {
            let ligneDate = convertirDateExcel(ligne['Date'] || '');
            if (ligneDate.includes('/')) { const p = ligneDate.split('/'); ligneDate = `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`; }
            if (filterDateDebut && ligneDate < filterDateDebut) match = false;
            if (filterDateFin && ligneDate > filterDateFin) match = false;
        }
        if (filterEquipe && ligne['Équipe'] !== filterEquipe) match = false;
        return match;
    });

    if (donneesFiltrees.length === 0) { alert('⚠️ Aucune donnée pour ces filtres.'); return; }
    afficherDashboard();
    alert(`✅ ${donneesFiltrees.length} ligne(s) affichée(s)`);
}

function reinitialiserFiltres() {
    document.getElementById('filterDateDebut').value = '';
    document.getElementById('filterDateFin').value = '';
    document.getElementById('filterEquipe').value = '';
    donneesFiltrees = [...donneesCompletes];
    appliquerFiltreEquipeParDefaut();
}

// ============================================
// EXPORT
// ============================================
function exporterCSV() {
    if (donneesFiltrees.length === 0) { alert('Aucune donnée à exporter'); return; }
    const headers = ['Date','Référence','Trigramme','Équipe','Heure','Prod Bonne','Rebuts','Équipement','Équipement Durée','Qualité','Qualité Durée','Organisation','Organisation Durée','Autres','Autres Durée','Commentaire'];
    let csvContent = headers.join(';') + '\n';
    donneesFiltrees.forEach(ligne => {
        let date = convertirDateExcel(ligne['Date'] || '');
        if (date) { try { const d = new Date(date); if (!isNaN(d)) date = `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`; } catch(e) {} }
        csvContent += [date, ligne['Référence']||'', ligne['Jour']||'', ligne['Équipe']||'', ligne['Heure']||'', ligne['Prod Bonne']||0, ligne['Rebuts']||0, ligne['Équipement']||'', ligne['Équipement Durée']||0, ligne['Qualité']||'', ligne['Qualité Durée']||0, ligne['Organisation']||'', ligne['Organisation Durée']||0, ligne['Autres']||'', ligne['Autres Durée']||0, ligne['Commentaire']||''].join(';') + '\n';
    });
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.setAttribute('href', URL.createObjectURL(blob));
    link.setAttribute('download', `HPH_CP17_Export_${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

async function exporterPDF() {
    const elementsToHide = [document.querySelector('.header-actions'), document.querySelector('.table-controls'), document.querySelector('.filters-section'), document.querySelector('.objectifs-section'), document.getElementById('loadingMessage'), document.getElementById('errorMessage')];
    const originalDisplayStates = elementsToHide.map(el => el ? el.style.display : '');
    elementsToHide.forEach(el => { if (el) el.style.display = 'none'; });
    await new Promise(resolve => setTimeout(resolve, 300));
    window.print();
    setTimeout(() => { elementsToHide.forEach((el, i) => { if (el) el.style.display = originalDisplayStates[i]; }); }, 1000);
}

// ============================================
// OBJECTIFS PAR RÉFÉRENCE
// ============================================
function chargerObjectifs() {
    const stored = localStorage.getItem('objectifs_hph_cp17');
    if (stored) {
        objectifs = JSON.parse(stored);
        document.getElementById('objectifProdEquipe').value = objectifs.prodEquipe;
        document.getElementById('objectifTauxRebuts').value = objectifs.tauxRebuts;
        document.getElementById('objectifTRS').value = objectifs.trs || 85;
    }
}

function appliquerObjectifs() {
    objectifs.prodEquipe = parseInt(document.getElementById('objectifProdEquipe').value) || 5000;
    objectifs.tauxRebuts = parseFloat(document.getElementById('objectifTauxRebuts').value) || 5;
    objectifs.trs = parseFloat(document.getElementById('objectifTRS').value) || 85;
    localStorage.setItem('objectifs_hph_cp17', JSON.stringify(objectifs));
    alert('✅ Objectifs enregistrés !');
    if (donneesFiltrees.length > 0) { calculerKPIs(); afficherGraphiques(); }
}

function rechercherObjectifsReference() {
    const reference = document.getElementById('referenceObjectif').value.trim().toUpperCase();
    if (!reference) return;
    const msg = document.getElementById('messageEnregistrement');
    msg.textContent = `ℹ️ Saisie manuelle pour ${reference}`;
    msg.style.color = '#2196F3';
    setTimeout(() => { msg.textContent = ''; }, 3000);
    if (donneesFiltrees.length > 0) afficherGraphiqueCadenceIoT();
}

function enregistrerObjectifsReference() {
    const reference = document.getElementById('referenceObjectif').value.trim().toUpperCase();
    if (!reference) { alert('⚠️ Veuillez saisir une référence.'); return; }
    const msg = document.getElementById('messageEnregistrement');
    msg.textContent = `✅ Objectifs prêts pour ${reference}.`;
    msg.style.color = '#4caf50';
    setTimeout(() => { msg.textContent = ''; }, 5000);
    appliquerObjectifs();
}

// ============================================
// INITIALISATION
// ============================================
document.addEventListener('DOMContentLoaded', function() {
    console.log('✅ Dashboard HPH CP17 IoT + MSAL initialisé');
    chargerObjectifs();
    chargerDonnees();
});
