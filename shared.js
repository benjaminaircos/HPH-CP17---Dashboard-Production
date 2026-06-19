// ============================================================
// HPH CP17 — shared.js
// Logique métier commune aux 3 pages (saisie / prod / manager)
// ============================================================

// ── URLs Power Automate ──────────────────────────────────────
const URL_LECTURE_SAISIE  = "https://defaultefdc051fe1144edaa7c5efec56da13.b1.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/68fb07af56e94845b714ce22d00b5f4c/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=hTGYXli6b3lv0VdHzQD8lKdAc07jyjerdzHloXKOV4c";
const URL_LECTURE_IOT     = "https://defaultefdc051fe1144edaa7c5efec56da13.b1.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/0e4d712f35204bbe9b246828013bf153/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=jan8mUdT2SOnljeXtb4iv98LBDPNKkQAKK76g9RyPRU";
const URL_ECRITURE_SAISIE = "https://defaultefdc051fe1144edaa7c5efec56da13.b1.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/5c714efc2639438580bc3cde147c9257/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=RXZFhVQGVJxlJhTyXYlB142AEO_TCb_4TT0iXdKJKCk";
const URL_EMPREINTES      = "https://defaultefdc051fe1144edaa7c5efec56da13.b1.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/3ed683137e644892b5d9f275c3438f52/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=kMUu5cvHflbBx-Lvo8C9b4TDmf8CClnw3YNm5oWQ1Hg";

// ── État global ──────────────────────────────────────────────
let donneesCompletes  = [];
let donneesFiltrees   = [];
let donneesIoT        = [];
let donneesEmpreintes = [];
let charts            = {};

let objectifs = {
    gamme:      6000,   // pcs/équipe (ex Objectif Prod Equipe)
    tauxRebuts: 4.0,
    trs:        85,
    cadence:    15,     // pcs/min nominale
    nbEmpreintes: 1
};

// ── Horaires équipes ─────────────────────────────────────────
function getHorairesEquipe(dateStr) {
    const jour = new Date(dateStr).getDay();
    if (jour === 1) return { Matin:[6,13],  Soir:[13,21], Nuit:[21,29] };
    if (jour === 5) return { Matin:[5,13],  Soir:[13,20], Nuit:[20,28] };
    return         { Matin:[5,13],  Soir:[13,21], Nuit:[21,29] };
}

function getDureeEquipe(equipe, dateStr) {
    const h = getHorairesEquipe(dateStr);
    if (!h[equipe]) return 480;
    return (h[equipe][1] - h[equipe][0]) * 60;
}

function determinerEquipeEnCours() {
    const heure = new Date().getHours();
    const jour  = new Date().getDay();
    let dM = 5, dS = 13, dN = 21;
    if (jour === 1) { dM = 6; }
    if (jour === 5) { dN = 20; }
    if (heure >= dM && heure < dS) return 'Matin';
    if (heure >= dS && heure < dN) return 'Soir';
    return 'Nuit';
}

// ── Conversion date Excel ────────────────────────────────────
function convertirDateExcel(valeur) {
    if (typeof valeur === 'string' && (valeur.includes('/') || valeur.includes('-'))) return valeur;
    let v = valeur;
    if (typeof valeur === 'string' && !isNaN(valeur) && valeur.trim() !== '') v = parseFloat(valeur);
    if (typeof v === 'number' && v > 1000) {
        const d = new Date(new Date(1899,11,30).getTime() + v * 86400000);
        return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    }
    return valeur;
}

function normaliserDate(val) {
    let d = convertirDateExcel(val || '');
    if (d.includes('/')) { const p = d.split('/'); d = `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`; }
    return d;
}

function dateAffichage(iso) {
    try {
        const d = new Date(iso);
        if (!isNaN(d)) return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
    } catch(e) {}
    return iso;
}

// ── Empreintes ───────────────────────────────────────────────
function getNbEmpreintes(reference) {
    if (!reference || !donneesEmpreintes.length) return objectifs.nbEmpreintes;
    const ref = donneesEmpreintes.find(r =>
        (r['Code article'] || r['Code_article'] || r['Code Article'] || '').toString().trim().toUpperCase()
        === reference.trim().toUpperCase()
    );
    if (ref) return parseInt(ref["nombre d'empreintes"] || ref['Nb_Empreintes'] || ref['empreintes'] || 1);
    return objectifs.nbEmpreintes;
}

// ── Calculs KPIs ─────────────────────────────────────────────
function calculerKPIsData(lignes) {
    let prodBonne = 0, rebuts = 0, arret = 0;
    lignes.forEach(l => {
        prodBonne += parseInt(l['Prod Bonne'] || l['Prod_x0020_Bonne'] || 0);
        rebuts    += parseInt(l['Rebuts'] || 0);
        arret     += parseInt(l['Équipement Durée'] || l['Équipement_x0020_Durée'] || 0)
                   + parseInt(l['Qualité Durée']     || l['Qualité_x0020_Durée']     || 0)
                   + parseInt(l['Organisation Durée'] || l['Organisation_x0020_Durée'] || 0)
                   + parseInt(l['Autres Durée']       || l['Autres_x0020_Durée']       || 0);
    });
    const total      = prodBonne + rebuts;
    const tauxRebuts = total > 0 ? (rebuts / total) * 100 : 0;
    const prodMax    = objectifs.cadence * 60 * lignes.length;
    const trs        = prodMax > 0 ? (total / prodMax) * 100 : 0;
    return { prodBonne, rebuts, total, tauxRebuts, arret, trs };
}

// ── Données par heure ────────────────────────────────────────
function agregerParHeure(lignes) {
    const map = {};
    lignes.forEach(l => {
        const h = l['Heure'] || '?';
        if (!map[h]) map[h] = { bonne:0, rebuts:0, arret:0 };
        map[h].bonne  += parseInt(l['Prod Bonne'] || l['Prod_x0020_Bonne'] || 0);
        map[h].rebuts += parseInt(l['Rebuts'] || 0);
        map[h].arret  += parseInt(l['Équipement Durée'] || 0)
                       + parseInt(l['Qualité Durée']    || 0)
                       + parseInt(l['Organisation Durée']|| 0)
                       + parseInt(l['Autres Durée']      || 0);
    });
    const heures = Object.keys(map).sort((a,b) => parseInt(a)-parseInt(b));
    return { map, heures };
}

// ── Cadence IoT ──────────────────────────────────────────────
function getCadenceIoTParHeure(dateISO, equipe) {
    if (!donneesIoT.length) return {};
    const hor = getHorairesEquipe(dateISO);
    const [debut] = hor[equipe] || [0,24];
    const map = {};
    donneesIoT.filter(ev => {
        return normaliserDate(ev['Date']) === dateISO
            && (ev['Etat']||'').toUpperCase() === 'CADENCE';
    }).forEach(ev => {
        const hEq = parseInt(ev['Heure']||0) - debut + 1;
        const lbl = `${hEq}h`;
        if (hEq >= 1 && hEq <= 8) {
            if (!map[lbl]) map[lbl] = [];
            map[lbl].push(parseFloat(ev['Cadence']||0) * objectifs.nbEmpreintes);
        }
    });
    const res = {};
    Object.keys(map).forEach(h => {
        const vals = map[h];
        res[h] = vals.reduce((a,b)=>a+b,0) / vals.length;
    });
    return res;
}

// ── Statut machine (bandeau IoT) ─────────────────────────────
function getDernierStatutMachine() {
    if (!donneesIoT.length) return null;
    const dernier = donneesIoT[donneesIoT.length - 1];
    const etat = (dernier['Etat']||'').toUpperCase();

    let cadence = null;
    for (let i = donneesIoT.length - 1; i >= 0; i--) {
        if ((donneesIoT[i]['Etat']||'').toUpperCase() === 'CADENCE') {
            cadence = parseFloat(donneesIoT[i]['Cadence']||0) * objectifs.nbEmpreintes;
            break;
        }
    }

    const heure  = parseInt(dernier['Heure']  || 0);
    const minute = parseInt(dernier['Minute'] || 0);
    const now    = new Date();
    const minDepuis = (now.getHours() * 60 + now.getMinutes()) - (heure * 60 + minute);

    return { etat, cadence, minDepuis: Math.max(0, minDepuis) };
}

function mettreAJourBandeauStatut() {
    const el = document.getElementById('statut-machine');
    if (!el) return;
    const s = getDernierStatutMachine();
    if (!s) { el.style.display = 'none'; return; }
    el.style.display = 'flex';

    const dot   = el.querySelector('.statut-dot');
    const label = el.querySelector('.statut-label');
    const since = el.querySelector('.statut-since');
    const cad   = el.querySelector('.statut-cadence');
    const tag   = el.querySelector('.refresh-tag');

    const marche = (s.etat === 'ALLUMAGE' || s.etat === 'ON');
    el.className = marche ? 'statut-marche' : 'statut-arret';
    label.textContent = marche ? 'En marche' : 'À l\'arrêt';
    since.textContent = `depuis ${s.minDepuis} min`;
    if (cad) cad.textContent = s.cadence
        ? (marche ? `Cadence : ${s.cadence.toFixed(1)} pcs/min` : `Dernière cadence : ${s.cadence.toFixed(1)} pcs/min`)
        : '';
    if (tag) tag.textContent = `IoT · maj il y a ${s.minDepuis} min`;
}

// ── Chargement données depuis Power Automate ─────────────────
async function chargerDonnees(urlSaisie, urlIoT, urlEmpreintes) {
    try {
        // Saisie
        const rSaisie = await fetch(urlSaisie, { method: 'GET' });
        if (!rSaisie.ok) throw new Error(`Saisie HTTP ${rSaisie.status}`);
        const dSaisie = await rSaisie.json();
        donneesCompletes = dSaisie.saisie || dSaisie.value || (Array.isArray(dSaisie) ? dSaisie : []);

        // IoT (optionnel)
        if (urlIoT) {
            try {
                const rIoT = await fetch(urlIoT, { method: 'GET' });
                if (rIoT.ok) { const d = await rIoT.json(); donneesIoT = d.value || d || []; }
            } catch(e) { console.warn('IoT non chargé:', e); }
        }

        // Empreintes (optionnel)
        if (urlEmpreintes) {
            try {
                const rEmp = await fetch(urlEmpreintes, { method: 'GET' });
                if (rEmp.ok) { const d = await rEmp.json(); donneesEmpreintes = d.value || d || []; }
            } catch(e) { console.warn('Empreintes non chargées:', e); }
        }

        console.log(`✅ Saisie: ${donneesCompletes.length} lignes | IoT: ${donneesIoT.length} | Empreintes: ${donneesEmpreintes.length}`);
        return true;
    } catch(e) {
        console.error('❌ Erreur chargement:', e);
        return false;
    }
}

// ── Filtrage données ─────────────────────────────────────────
function filtrerDonnees(dateDebut, dateFin, equipe) {
    return donneesCompletes.filter(l => {
        const d = normaliserDate(l['Date']);
        if (dateDebut && d < dateDebut) return false;
        if (dateFin   && d > dateFin)   return false;
        if (equipe && l['Équipe'] !== equipe) return false;
        return true;
    });
}

function filtrerAujourdhui() {
    const today  = new Date().toISOString().split('T')[0];
    const equipe = determinerEquipeEnCours();
    return filtrerDonnees(today, today, equipe);
}

// ── Horloge ──────────────────────────────────────────────────
function startClock(elId) {
    const el = document.getElementById(elId);
    if (!el) return;
    function tick() {
        const n = new Date();
        el.textContent = n.toLocaleDateString('fr-FR') + ' — ' + n.toLocaleTimeString('fr-FR');
    }
    tick();
    setInterval(tick, 1000);
}

// ── Chart helpers ─────────────────────────────────────────────
const CHART_COLORS = {
    ok:     '#2dd4a0', warn: '#f0b429', danger: '#e8502a',
    blue:   '#378add', purple:'#a78bfa',
    grid:   'rgba(42,50,69,.7)', tick: '#7a8ba8'
};

function baseChartOpts(yLabel, yMax) {
    return {
        responsive: true, maintainAspectRatio: true,
        plugins: { legend: { display: true, labels: { color: CHART_COLORS.tick, font:{size:10}, boxWidth:12 } } },
        scales: {
            x: { grid:{color:CHART_COLORS.grid}, ticks:{color:CHART_COLORS.tick, font:{size:11}} },
            y: {
                grid:{color:CHART_COLORS.grid}, ticks:{color:CHART_COLORS.tick, font:{size:11}},
                beginAtZero: true,
                ...(yMax ? {max:yMax} : {}),
                ...(yLabel ? {title:{display:true, text:yLabel, color:CHART_COLORS.tick, font:{size:10}}} : {})
            }
        }
    };
}

function makeOrUpdate(key, canvasId, config) {
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;
    if (charts[key]) charts[key].destroy();
    charts[key] = new Chart(ctx, config);
}

// ── Graphique Production par heure ───────────────────────────
function chartProductionHeure(canvasId, heures, prodBonne, gamme) {
    const objectifH = gamme / 8;
    const colors = prodBonne.map(v => v >= objectifH ? CHART_COLORS.ok+'bb' : CHART_COLORS.danger+'bb');
    makeOrUpdate('prodHeure', canvasId, {
        type: 'bar',
        data: { labels: heures, datasets: [
            { label:'Production Bonne (pcs)', data:prodBonne, backgroundColor:colors, borderColor:colors.map(c=>c.replace('bb','ff')), borderWidth:1, borderRadius:3, order:2 },
            { label:`Gamme planning (${objectifH.toFixed(0)} pcs/h)`, data:heures.map(()=>objectifH), type:'line', borderColor:CHART_COLORS.warn, borderDash:[8,4], borderWidth:2, pointRadius:0, fill:false, order:1 }
        ]},
        options: { ...baseChartOpts('Production (pcs)'), plugins:{ legend:{display:true, labels:{color:CHART_COLORS.tick, font:{size:10}, boxWidth:12}} } }
    });
}

// ── Graphique TRS par heure ───────────────────────────────────
function chartTRSHeure(canvasId, heures, trsData, objTRS) {
    const colors = trsData.map(v => v !== null && v >= objTRS ? CHART_COLORS.blue+'bb' : CHART_COLORS.danger+'bb');
    makeOrUpdate('trsHeure', canvasId, {
        type: 'bar',
        data: { labels: heures, datasets: [
            { label:'TRS Global (%)', data:trsData, backgroundColor:colors, borderColor:colors.map(c=>c.replace('bb','ff')), borderWidth:1, borderRadius:3, order:2 },
            { label:`Objectif TRS (${objTRS}%)`, data:heures.map(()=>objTRS), type:'line', borderColor:CHART_COLORS.warn, borderDash:[8,4], borderWidth:2, pointRadius:0, fill:false, order:1 }
        ]},
        options: { ...baseChartOpts('TRS (%)', 170) }
    });
}

// ── Graphique Temps non justifié ──────────────────────────────
function chartTempsNonJustifie(canvasId, heures, data) {
    const colors = data.map(v => v === 0 ? CHART_COLORS.ok+'bb' : v <= 10 ? CHART_COLORS.warn+'bb' : CHART_COLORS.danger+'bb');
    makeOrUpdate('tempsNJ', canvasId, {
        type: 'bar',
        data: { labels: heures, datasets: [
            { label:'Temps non justifié (min)', data, backgroundColor:colors, borderColor:colors.map(c=>c.replace('bb','ff')), borderWidth:1, borderRadius:3 }
        ]},
        options: { ...baseChartOpts('Durée (min)'), plugins:{legend:{display:false}} }
    });
}

// ── Graphique Taux de performance ─────────────────────────────
function chartPerformance(canvasId, heures, perfData) {
    const colors = perfData.map(v => v === null ? 'transparent' : v >= 100 ? CHART_COLORS.ok+'bb' : v >= 85 ? CHART_COLORS.warn+'bb' : CHART_COLORS.danger+'bb');
    makeOrUpdate('performance', canvasId, {
        type: 'bar',
        data: { labels: heures, datasets: [
            { label:'Taux de performance (%)', data:perfData, backgroundColor:colors, borderColor:colors.map(c=>c.replace('bb','ff')), borderWidth:1, borderRadius:3, order:2, spanGaps:true },
            { label:'Objectif (100%)', data:heures.map(()=>100), type:'line', borderColor:CHART_COLORS.warn, borderDash:[8,4], borderWidth:2, pointRadius:0, fill:false, order:1 }
        ]},
        options: { ...baseChartOpts('Performance (%)', 170) }
    });
}

// ── Graphique Cadence IoT (intervalles variables) ─────────────
function chartCadenceIoT(canvasId, dateISO, equipe, intervalMin) {
    const hor    = getHorairesEquipe(dateISO);
    const [deb, fin] = hor[equipe] || [13, 21];
    const labels = [], cadences = [];
    const nb     = Math.floor((fin - deb) * 60 / intervalMin);

    for (let i = 0; i < nb; i++) {
        const t   = deb * 60 + i * intervalMin;
        const hh  = String(Math.floor(t/60)).padStart(2,'0');
        const mm  = String(t%60).padStart(2,'0');
        labels.push(`${hh}h${mm}`);

        // Chercher les events CADENCE IoT dans cet intervalle
        const tFin = t + intervalMin;
        const evts = donneesIoT.filter(ev => {
            if (normaliserDate(ev['Date']) !== dateISO) return false;
            if ((ev['Etat']||'').toUpperCase() !== 'CADENCE') return false;
            const evMin = parseInt(ev['Heure']||0)*60 + parseInt(ev['Minute']||0);
            return evMin >= t && evMin < tFin;
        });
        if (evts.length > 0) {
            const avg = evts.reduce((s,e) => s + parseFloat(e['Cadence']||0), 0) / evts.length;
            cadences.push(+(avg * objectifs.nbEmpreintes).toFixed(2));
        } else {
            cadences.push(null);
        }
    }

    makeOrUpdate('cadenceIoT', canvasId, {
        type: 'line',
        data: { labels, datasets: [
            { label:'Cadence réelle (pcs/min)', data:cadences, borderColor:CHART_COLORS.blue, backgroundColor:CHART_COLORS.blue+'22', borderWidth:2, pointRadius:3, fill:true, tension:0.3, spanGaps:true },
            { label:`Cadence nominale (${objectifs.cadence} pcs/min)`, data:labels.map(()=>objectifs.cadence), borderColor:CHART_COLORS.warn, borderDash:[8,4], borderWidth:1.5, pointRadius:0, fill:false }
        ]},
        options: {
            ...baseChartOpts('pcs/min'),
            plugins:{ legend:{display:true, labels:{color:CHART_COLORS.tick, font:{size:10}, boxWidth:12}} },
            scales: {
                x: { grid:{color:CHART_COLORS.grid}, ticks:{color:CHART_COLORS.tick, font:{size:10}, maxRotation:45} },
                y: { grid:{color:CHART_COLORS.grid}, ticks:{color:CHART_COLORS.tick, font:{size:11}}, beginAtZero:false, title:{display:true, text:'pcs/min', color:CHART_COLORS.tick, font:{size:10}} }
            }
        }
    });
}
