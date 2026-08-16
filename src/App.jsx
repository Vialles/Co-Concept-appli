import React, { useState, useMemo, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { auth, db, storage } from "./firebase";
import {
  onAuthStateChanged,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  signOut,
} from "firebase/auth";
import { doc, getDoc, setDoc, onSnapshot, collection, addDoc, deleteDoc } from "firebase/firestore";
import { ref as storageRef, uploadBytes, getDownloadURL, listAll, deleteObject } from "firebase/storage";
import {
  Radar,
  FileText,
  KanbanSquare,
  MapPin,
  Calendar,
  Euro,
  Search,
  Plus,
  Check,
  Clock,
  Building2,
  Trophy,
  X,
  Globe,
  ExternalLink,
  AlertTriangle,
  Users,
  Download,
  CalendarDays,
  ChevronLeft,
  ChevronRight,
  Wallet,
  FileSpreadsheet,
  LogOut,
  Loader2,
  Pencil,
  LayoutDashboard,
  Trash2,
  Archive,
  Image,
  Upload,
} from "lucide-react";

/* ---------------------------------------------------------
   TOKENS
   Paper: #EEF0EA (grey-green paper, not warm cream)
   Ink:   #202B21 (deep green-black, technical-drawing ink)
   Moss:  #4C6444 (primary accent — vegetation)
   Clay:  #B07A46 (secondary accent — gravel path)
   Blue:  #46618C (blueprint line — data / meta)
   Rust:  #9C4B3B (alert — deadlines)
--------------------------------------------------------- */

const TOKENS = {
  paper: "#EEF0EA",
  paperDim: "#E4E7DD",
  ink: "#202B21",
  inkSoft: "#5A6357",
  moss: "#4C6444",
  mossDim: "#DDE4D6",
  clay: "#B07A46",
  clayDim: "#F0E3D2",
  blue: "#46618C",
  blueDim: "#DEE5EF",
  rust: "#9C4B3B",
  rustDim: "#F1DFDA",
  line: "#C8CDBF",
};

const FONT_IMPORT = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');
`;

/* ---------------------------------------------------------
   MOCK DATA
--------------------------------------------------------- */

const MOCK_AO = [
  {
    id: "AO-2607",
    titre: "Aménagement paysager du parc de la Presqu'île",
    acheteur: "Communauté d'agglomération de La Rochelle",
    zone: "La Rochelle (17)",
    dateLimit: "2026-08-22",
    montant: "180 000 – 240 000 €",
    tags: ["espaces verts", "conception"],
    type: "conception",
    nouveau: true,
  },
  {
    id: "AO-2591",
    titre: "Entretien des espaces verts communaux — marché à bons de commande",
    acheteur: "Mairie de Rochefort",
    zone: "Rochefort (17)",
    dateLimit: "2026-08-30",
    montant: "90 000 €/an",
    tags: ["entretien", "espaces verts"],
    type: "chantier",
    nouveau: true,
  },
  {
    id: "AO-2544",
    titre: "Conception et réalisation d'un jardin pédagogique",
    acheteur: "Département de Charente-Maritime",
    zone: "Saintes (17)",
    dateLimit: "2026-09-05",
    montant: "65 000 €",
    tags: ["conception", "aménagement", "biodiversité"],
    type: "conception",
    nouveau: false,
  },
  {
    id: "AO-2498",
    titre: "Requalification paysagère d'une friche urbaine",
    acheteur: "Ville de Niort",
    zone: "Niort (79)",
    dateLimit: "2026-09-18",
    montant: "310 000 €",
    tags: ["aménagement", "conception", "génie écologique"],
    type: "conception",
    nouveau: false,
  },
  {
    id: "AO-STRAM-3",
    titre: "Am Stram Graines #3 — Aménagement et végétalisation de cours d'école",
    acheteur: "CAUE 17 (Conseil d'Architecture, d'Urbanisme et de l'Environnement de la Charente-Maritime)",
    zone: "Charente-Maritime (17) — communes du département, alentours La Rochelle",
    dateLimit: null,
    montant: null,
    tags: ["conception", "aménagement", "cour d'école", "végétalisation", "participatif"],
    type: "conception",
    nouveau: true,
    sourceUrl: "https://www.caue17.fr",
    infoAVerifier: true,
    note: "Appel à candidatures coordonné par le CAUE 17 pour désigner les équipes de maîtrise d'œuvre chargées de projets de végétalisation de cours d'école en Charente-Maritime. Date limite et montant de l'édition #3 non publiés sur les pages consultées — à vérifier directement sur caue17.fr.",
  },
];

const DEFAULT_PLATFORMS = [
  {
    id: "boamp",
    name: "BOAMP — Bulletin officiel des annonces de marchés publics",
    url: "https://www.boamp.fr",
    active: true,
  },
  {
    id: "place",
    name: "PLACE — Plateforme des achats de l'État",
    url: "https://www.marches-publics.gouv.fr",
    active: true,
  },
  {
    id: "caue17",
    name: "CAUE 17 — Appels à candidatures et à projets",
    url: "https://www.caue17.fr",
    active: true,
  },
  {
    id: "aws-achat",
    name: "AWS-Achat — profil acheteur mutualisé (national)",
    url: "https://www.marches-publics.info",
    active: true,
  },
  {
    id: "aws-aquitaine",
    name: "AWS-Achat — portail régional Aquitaine",
    url: "http://aquitaine.marches-publics.info",
    active: true,
  },
  {
    id: "demat-ampa",
    name: "Profil acheteur Région Nouvelle-Aquitaine (AMPA)",
    url: "https://demat-ampa.fr",
    active: true,
  },
  {
    id: "e-marchespublics",
    name: "e-marchespublics.com — profil acheteur et dépôt d'offres",
    url: "https://www.e-marchespublics.com",
    active: false,
  },
  {
    id: "marches-securises",
    name: "Marchés Sécurisés",
    url: "https://www.marches-securises.fr",
    active: false,
  },
  {
    id: "urcaue",
    name: "URCAUE Nouvelle-Aquitaine",
    url: "http://www.urcaue-na.fr",
    active: false,
  },
  {
    id: "ffp",
    name: "Fédération Française du Paysage — Actualités",
    url: "https://f-f-p.org",
    active: false,
  },
];

const TYPE_LABEL = {
  conception: "Conception",
  chantier: "Travaux / Entretien",
};

const PIECE_STATUTS = [
  { value: "a_faire", label: "À faire" },
  { value: "en_cours", label: "En cours" },
  { value: "termine", label: "Terminé" },
  { value: "pas_demande", label: "Pas demandé" },
];

const PIECES_ADMIN = [
  "Kbis de moins de 3 mois",
  "Attestation d'assurance décennale",
  "Attestation de régularité fiscale et sociale",
  "DC1 — Lettre de candidature",
  "DC2 — Déclaration du candidat",
];

// Pièces spécifiques au dossier, selon que l'AO porte sur des travaux/entretien ou de la conception
const PIECES_PAR_TYPE = {
  chantier: [
    "Mémoire technique",
    "Références de chantiers similaires",
    "Fiches techniques du matériel utilisé",
  ],
  conception: [
    "Mémoire technique",
    "Références de projets comparables (book)",
    "Plan de masse / plan d'aménagement",
    "Coupes et élévations",
    "Perspective ou insertion paysagère",
    "Notice paysagère",
    "Palette végétale",
  ],
};

const MEMO_SECTIONS_PAR_TYPE = {
  chantier: [
    { key: "presentation", label: "Présentation de l'entreprise" },
    { key: "moyens", label: "Moyens humains et matériels" },
    { key: "methodologie", label: "Méthodologie d'intervention" },
    { key: "references", label: "Références de chantiers comparables" },
    { key: "planning", label: "Planning prévisionnel d'intervention" },
  ],
  conception: [
    { key: "presentation", label: "Présentation de l'entreprise et de l'équipe de maîtrise d'œuvre" },
    { key: "partiPris", label: "Parti pris paysager et démarche de conception" },
    { key: "references", label: "Références visuelles et projets comparables" },
    { key: "graphique", label: "Description des pièces graphiques produites" },
    { key: "planning", label: "Planning prévisionnel de conception et de réalisation" },
  ],
};

const STATUTS = ["Repéré", "En préparation", "Déposé", "Résultat"];

const MEMBER_COLORS = [TOKENS.moss, TOKENS.blue, TOKENS.clay, TOKENS.rust];
const WEEKDAY_LETTERS = ["L", "M", "M", "J", "V", "S", "D"];

// Fiches projet détaillées pour quelques AO suivis (équipe, planning, réunions)
const PROJECT_DETAILS = {
  "AO-2591": {
    team: [
      { id: "m1", name: "Camille B.", poste: "Cheffe de projet", mandataire: true },
      { id: "m2", name: "Nicolas T.", poste: "Chargé d'affaires", mandataire: false },
      { id: "m3", name: "Sofia R.", poste: "Chef d'équipe entretien", mandataire: false },
    ],
    tasks: [
      { id: "t1", label: "Analyse du cahier des charges", assigneeIds: ["m1"], start: "2026-08-07", end: "2026-08-12" },
      { id: "t2", label: "Chiffrage et devis matériel", assigneeIds: ["m2"], start: "2026-08-11", end: "2026-08-18" },
      { id: "t3", label: "Rédaction mémoire technique", assigneeIds: ["m1"], start: "2026-08-13", end: "2026-08-22" },
      { id: "t4", label: "Validation planning d'intervention", assigneeIds: ["m3"], start: "2026-08-19", end: "2026-08-25" },
      { id: "t5", label: "Dépôt du dossier", assigneeIds: ["m1"], start: "2026-08-26", end: "2026-08-29" },
    ],
    meetings: [
      { id: "r1", title: "Cadrage interne du dossier", date: "2026-08-08", time: "09:30", attendees: ["m1", "m2"] },
      { id: "r2", title: "Point chiffrage avec Sofia", date: "2026-08-18", time: "14:00", attendees: ["m1", "m2", "m3"] },
      { id: "r3", title: "Relecture finale avant dépôt", date: "2026-08-27", time: "10:00", attendees: ["m1", "m2", "m3"] },
    ],
    finance: { caByMonth: { "2026-08": "8 400" }, hours: { m1: { "2026-08": 14 }, m2: { "2026-08": 9 }, m3: { "2026-08": 6 } } },
  },
};

function defaultPieces(type) {
  return {
    piecesAdmin: PIECES_ADMIN.map((label, i) => ({ id: "a" + i, label, status: "a_faire" })),
    piecesTech: PIECES_PAR_TYPE[type].map((label, i) => ({ id: "t" + i, label, status: "a_faire" })),
  };
}

function defaultProjectDetail(ao, followedDate) {
  const start = new Date(followedDate ?? "2026-08-06");
  const end = ao.dateLimit ? new Date(ao.dateLimit) : new Date(start.getTime() + 60 * 86400000);
  const mid = new Date((start.getTime() + end.getTime()) / 2);
  const fmt = (d) => d.toISOString().slice(0, 10);
  const isConception = ao.type === "conception";
  return {
    team: [
      { id: "d1", name: "", poste: "", mandataire: true },
      { id: "d2", name: "", poste: isConception ? "Paysagiste concepteur" : "Chargé d'affaires", mandataire: false },
    ],
    tasks: [
      { id: "d-t1", label: "Analyse du cahier des charges", assigneeIds: ["d1"], start: fmt(start), end: fmt(new Date(start.getTime() + 4 * 86400000)) },
      { id: "d-t2", label: isConception ? "Esquisses et pièces graphiques" : "Chiffrage et moyens", assigneeIds: ["d2"], start: fmt(new Date(start.getTime() + 3 * 86400000)), end: fmt(mid) },
      { id: "d-t3", label: "Rédaction du dossier", assigneeIds: ["d1"], start: fmt(mid), end: fmt(new Date(end.getTime() - 3 * 86400000)) },
      { id: "d-t4", label: "Dépôt", assigneeIds: ["d1"], start: fmt(new Date(end.getTime() - 2 * 86400000)), end: fmt(end) },
    ],
    meetings: [
      { id: "d-r1", title: "Cadrage interne", date: fmt(new Date(start.getTime() + 1 * 86400000)), time: "09:30", attendees: ["d1", "d2"] },
    ],
    links: { canva: "", onenote: "" },
    finance: { caByMonth: {}, hours: {} },
    ...defaultPieces(ao.type),
    memo: {},
    rangeStart: fmt(start),
    rangeEnd: fmt(end),
  };
}

function getProjectDetail(ao, followedDate) {
  const base = PROJECT_DETAILS[ao.id];
  if (base) {
    return {
      ...base,
      links: base.links ?? { canva: "", onenote: "" },
      finance: base.finance ?? { caByMonth: {}, hours: {} },
      piecesAdmin: base.piecesAdmin ?? defaultPieces(ao.type).piecesAdmin,
      piecesTech: base.piecesTech ?? defaultPieces(ao.type).piecesTech,
      memo: base.memo ?? {},
      rangeStart: followedDate ?? base.tasks[0]?.start,
      rangeEnd: ao.dateLimit ?? base.tasks[base.tasks.length - 1]?.end,
    };
  }
  return defaultProjectDetail(ao, followedDate);
}

// Complète les champs manquants sur une fiche projet déjà existante dans
// Firestore (créée avant l'ajout des pièces / de la mémoire technique, par
// exemple) — évite les plantages sur d.piecesAdmin/d.piecesTech undefined.
function withProjectDefaults(ao, data) {
  const defaults = defaultPieces(ao.type);
  return {
    team: data.team ?? [],
    tasks: (data.tasks ?? []).map((t) => ({
      ...t,
      assigneeIds: t.assigneeIds ?? (t.assigneeId ? [t.assigneeId] : []),
    })),
    meetings: data.meetings ?? [],
    links: data.links ?? { canva: "", onenote: "" },
    finance: data.finance ?? { caByMonth: {}, hours: {} },
    piecesAdmin: data.piecesAdmin ?? defaults.piecesAdmin,
    piecesTech: data.piecesTech ?? defaults.piecesTech,
    memo: data.memo ?? {},
    rangeStart: data.rangeStart,
    rangeEnd: data.rangeEnd,
  };
}

function buildICS(events) {
  const dt = (dateStr, timeStr) => dateStr.replace(/-/g, "") + "T" + timeStr.replace(":", "") + "00";
  const lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//AO Paysage//FR"];
  events.forEach((e) => {
    lines.push(
      "BEGIN:VEVENT",
      "UID:" + e.id + "@ao-paysage",
      "SUMMARY:" + e.title,
      "DTSTART:" + dt(e.date, e.time),
      "DTEND:" + dt(e.date, e.time),
      "END:VEVENT"
    );
  });
  lines.push("END:VCALENDAR");
  return lines.join("\r\n");
}

function downloadICS(filename, content) {
  const blob = new Blob([content], { type: "text/calendar" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

// Insère un événement (créneau horaire précis) dans le Google Agenda de la
// personne connectée. Nécessite un token d'accès obtenu via Google Identity
// Services (connexion à la demande, valable ~1h).
async function pushEventToGoogle(token, { title, description, dateStr, time, durationMinutes = 60 }) {
  const start = new Date(`${dateStr}T${time}:00`);
  const end = new Date(start.getTime() + durationMinutes * 60000);
  const res = await fetch("https://www.googleapis.com/calendar/v3/calendars/primary/events", {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({
      summary: title,
      description: description ?? "",
      start: { dateTime: start.toISOString(), timeZone: "Europe/Paris" },
      end: { dateTime: end.toISOString(), timeZone: "Europe/Paris" },
    }),
  });
  if (!res.ok) throw new Error((await res.json()).error?.message ?? "Erreur Google Agenda");
  return res.json();
}

// Insère un événement journée entière (ex. date limite de dépôt) dans le
// Google Agenda de la personne connectée.
async function pushAllDayEventToGoogle(token, { title, description, dateStr }) {
  const end = new Date(dateStr);
  end.setDate(end.getDate() + 1);
  const res = await fetch("https://www.googleapis.com/calendar/v3/calendars/primary/events", {
    method: "POST",
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: JSON.stringify({
      summary: title,
      description: description ?? "",
      start: { date: dateStr },
      end: { date: end.toISOString().slice(0, 10) },
    }),
  });
  if (!res.ok) throw new Error((await res.json()).error?.message ?? "Erreur Google Agenda");
  return res.json();
}

/* ---------------------------------------------------------
   SMALL PRIMITIVES
--------------------------------------------------------- */

function Cartouche({ activeTab }) {
  const tabLabel = {
    veille: "Veille",
    redaction: "Rédaction",
    suivi: "Suivi",
    finances: "Suivi financier",
    sources: "Sources",
    ensemble: "Vue d'ensemble",
    historique: "Historique",
  }[activeTab];

  return (
    <div
      className="border-b"
      style={{ borderColor: TOKENS.ink, background: TOKENS.paper }}
    >
      <div className="grid grid-cols-12">
        <div className="col-span-8 sm:col-span-9 px-6 py-4 border-r flex items-center gap-4" style={{ borderColor: TOKENS.line }}>
          <img src="/logo-co-concept.png" alt="Co-Concept" style={{ height: 44, width: "auto" }} />
          <div>
            <div
              className="text-xs tracking-widest uppercase mb-1"
              style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace", letterSpacing: "0.15em" }}
            >
              Outil de réponse aux appels d'offres
            </div>
            <h1
              className="text-2xl sm:text-3xl"
              style={{ color: TOKENS.ink, fontFamily: "'Fraunces', serif", fontWeight: 600 }}
            >
              AO&nbsp;Paysage
            </h1>
          </div>
        </div>
        <div className="col-span-4 sm:col-span-3 px-4 py-4 flex flex-col justify-center">
          <div
            className="text-[10px] uppercase tracking-wider mb-1"
            style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}
          >
            Module
          </div>
          <div
            className="text-sm"
            style={{ color: TOKENS.ink, fontFamily: "'JetBrains Mono', monospace", fontWeight: 500 }}
          >
            {tabLabel}
          </div>
          <div
            className="text-[10px] uppercase tracking-wider mt-2 mb-1"
            style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}
          >
            Indice
          </div>
          <div
            className="text-sm"
            style={{ color: TOKENS.ink, fontFamily: "'JetBrains Mono', monospace" }}
          >
            Prototype — A
          </div>
        </div>
      </div>
    </div>
  );
}

function NavItem({ icon: Icon, label, active, onClick }) {
  return (
    <button
      onClick={onClick}
      className="w-full flex items-center gap-3 px-4 py-3 text-left transition-colors"
      style={{
        background: active ? TOKENS.mossDim : "transparent",
        color: active ? TOKENS.moss : TOKENS.inkSoft,
        borderLeft: `3px solid ${active ? TOKENS.moss : "transparent"}`,
      }}
    >
      <Icon size={17} strokeWidth={2} />
      <span
        className="text-sm"
        style={{ fontFamily: "'Inter', sans-serif", fontWeight: active ? 600 : 500 }}
      >
        {label}
      </span>
    </button>
  );
}

function Tag({ children, tone = "moss" }) {
  const map = {
    moss: { bg: TOKENS.mossDim, fg: TOKENS.moss },
    clay: { bg: TOKENS.clayDim, fg: TOKENS.clay },
    blue: { bg: TOKENS.blueDim, fg: TOKENS.blue },
    rust: { bg: TOKENS.rustDim, fg: TOKENS.rust },
  }[tone];
  return (
    <span
      className="text-[11px] px-2 py-0.5 rounded-full"
      style={{
        background: map.bg,
        color: map.fg,
        fontFamily: "'JetBrains Mono', monospace",
      }}
    >
      {children}
    </span>
  );
}

function daysLeft(dateStr) {
  if (!dateStr) return null;
  const diff = Math.ceil(
    (new Date(dateStr) - new Date("2026-08-06")) / (1000 * 60 * 60 * 24)
  );
  return diff;
}

/* ---------------------------------------------------------
   VEILLE
--------------------------------------------------------- */

function AddAOForm({ platforms, initial, onCancel, onSubmit }) {
  const [titre, setTitre] = useState(initial?.titre ?? "");
  const [acheteur, setAcheteur] = useState(initial?.acheteur ?? "");
  const [zone, setZone] = useState(initial?.zone ?? "");
  const [dateLimit, setDateLimit] = useState(initial?.dateLimit ?? "");
  const [montant, setMontant] = useState(initial?.montant ?? "");
  const [type, setType] = useState(initial?.type ?? "conception");
  const [tags, setTags] = useState((initial?.tags ?? []).join(", "));
  const [sourceId, setSourceId] = useState(
    platforms.find((p) => p.name === initial?.sourceName)?.id ?? ""
  );
  const [dateDebut, setDateDebut] = useState(initial?.dateDebut ?? "");
  const [dateFin, setDateFin] = useState(initial?.dateFin ?? "");
  const [dateResultat, setDateResultat] = useState(initial?.dateResultat ?? "");
  const [resultat, setResultat] = useState(initial?.resultat ?? "en_attente");

  const inputStyle = {
    borderColor: TOKENS.line,
    fontFamily: "'Inter', sans-serif",
    color: TOKENS.ink,
    background: "white",
  };

  const submit = (e) => {
    e.preventDefault();
    if (!titre.trim()) return;
    onSubmit({
      titre: titre.trim(),
      acheteur: acheteur.trim(),
      zone: zone.trim(),
      dateLimit: dateLimit || null,
      montant: montant.trim() || null,
      type,
      tags: tags
        .split(",")
        .map((t) => t.trim())
        .filter(Boolean),
      nouveau: initial?.nouveau ?? true,
      custom: true,
      sourceName: platforms.find((p) => p.id === sourceId)?.name ?? "",
      dateDebut: dateDebut || null,
      dateFin: dateFin || null,
      dateResultat: dateResultat || null,
      resultat,
    });
  };

  return (
    <form onSubmit={submit} className="mb-6 p-4 border flex flex-col gap-2" style={{ borderColor: TOKENS.moss, background: "white" }}>
      <div className="grid sm:grid-cols-2 gap-2">
        <input required value={titre} onChange={(e) => setTitre(e.target.value)} placeholder="Titre de l'AO" className="p-2 text-sm border outline-none sm:col-span-2" style={inputStyle} />
        <input value={acheteur} onChange={(e) => setAcheteur(e.target.value)} placeholder="Maître d'ouvrage / acheteur" className="p-2 text-sm border outline-none" style={inputStyle} />
        <input value={zone} onChange={(e) => setZone(e.target.value)} placeholder="Zone / commune" className="p-2 text-sm border outline-none" style={inputStyle} />
        <div>
          <label className="text-[10px] uppercase tracking-wider block mb-0.5" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>Date limite de dépôt</label>
          <input type="date" value={dateLimit} onChange={(e) => setDateLimit(e.target.value)} className="p-2 text-sm border outline-none w-full" style={inputStyle} />
        </div>
        <input value={montant} onChange={(e) => setMontant(e.target.value)} placeholder="Montant estimé" className="p-2 text-sm border outline-none" style={inputStyle} />
        <select value={type} onChange={(e) => setType(e.target.value)} className="p-2 text-sm border outline-none" style={inputStyle}>
          <option value="conception">Conception</option>
          <option value="chantier">Travaux / Entretien</option>
        </select>
        <select value={sourceId} onChange={(e) => setSourceId(e.target.value)} className="p-2 text-sm border outline-none" style={inputStyle}>
          <option value="">Trouvé sur… (optionnel)</option>
          {platforms.map((p) => (
            <option key={p.id} value={p.id}>{p.name}</option>
          ))}
        </select>
        <input value={tags} onChange={(e) => setTags(e.target.value)} placeholder="Mots-clés séparés par des virgules" className="p-2 text-sm border outline-none sm:col-span-2" style={inputStyle} />
      </div>

      <div className="mt-2 pt-2 border-t" style={{ borderColor: TOKENS.line }}>
        <div className="text-[10px] uppercase tracking-wider mb-2" style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}>
          Pour la vue d'ensemble (planning et suivi financier)
        </div>
        <div className="grid sm:grid-cols-2 gap-2">
          <div>
            <label className="text-[10px] uppercase tracking-wider block mb-0.5" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>Début du projet</label>
            <input type="date" value={dateDebut} onChange={(e) => setDateDebut(e.target.value)} className="p-2 text-sm border outline-none w-full" style={inputStyle} />
          </div>
          <div>
            <label className="text-[10px] uppercase tracking-wider block mb-0.5" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>Fin du projet</label>
            <input type="date" value={dateFin} onChange={(e) => setDateFin(e.target.value)} className="p-2 text-sm border outline-none w-full" style={inputStyle} />
          </div>
          <div>
            <label className="text-[10px] uppercase tracking-wider block mb-0.5" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>Date de résultat</label>
            <input type="date" value={dateResultat} onChange={(e) => setDateResultat(e.target.value)} className="p-2 text-sm border outline-none w-full" style={inputStyle} />
          </div>
          <div>
            <label className="text-[10px] uppercase tracking-wider block mb-0.5" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>Résultat</label>
            <select value={resultat} onChange={(e) => setResultat(e.target.value)} className="p-2 text-sm border outline-none w-full" style={inputStyle}>
              <option value="en_attente">En attente</option>
              <option value="gagne">Gagné</option>
              <option value="perdu">Perdu</option>
            </select>
          </div>
        </div>
      </div>

      <div className="flex gap-2 mt-1">
        <button type="submit" className="text-sm px-3 py-2" style={{ background: TOKENS.ink, color: TOKENS.paper, fontFamily: "'Inter', sans-serif", fontWeight: 600 }}>
          {initial ? "Enregistrer les modifications" : "Ajouter cet AO"}
        </button>
        <button type="button" onClick={onCancel} className="text-sm px-3 py-2" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
          Annuler
        </button>
      </div>
    </form>
  );
}

function Veille({ aos, followed, onFollow, onAddAO, onUpdateAO, onDeleteAO, onArchiveAO, platforms, googleConnected, googleToken, connectGoogle }) {
  const [query, setQuery] = useState("");
  const [sendingId, setSendingId] = useState(null);
  const [sentIds, setSentIds] = useState({});

  const sendDeadline = async (ao) => {
    if (!googleConnected) {
      connectGoogle();
      return;
    }
    setSendingId(ao.id);
    try {
      await pushAllDayEventToGoogle(googleToken, {
        title: "Date limite de dépôt — " + ao.titre,
        description: (ao.acheteur ?? "") + (ao.zone ? " — " + ao.zone : ""),
        dateStr: ao.dateLimit,
      });
      setSentIds((s) => ({ ...s, [ao.id]: true }));
    } catch (err) {
      alert("Échec de l'envoi vers Google Agenda : " + err.message);
    } finally {
      setSendingId(null);
    }
  };
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const filtered = useMemo(
    () =>
      aos.filter((a) =>
        (a.titre + a.zone + a.tags.join(" ")).toLowerCase().includes(query.toLowerCase())
      ),
    [aos, query]
  );
  const editingAO = aos.find((a) => a.id === editingId);

  return (
    <div className="p-6">
      <div className="flex items-center gap-3 mb-4">
        <div
          className="flex items-center gap-2 px-3 py-2 flex-1 max-w-md border"
          style={{ borderColor: TOKENS.line, background: "white" }}
        >
          <Search size={15} color={TOKENS.inkSoft} />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Mot-clé, commune, type de marché…"
            className="flex-1 outline-none text-sm bg-transparent"
            style={{ fontFamily: "'Inter', sans-serif", color: TOKENS.ink }}
          />
        </div>
        <div
          className="text-xs px-3 py-2"
          style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}
        >
          {filtered.length} appel{filtered.length > 1 ? "s" : ""} d'offres repéré{filtered.length > 1 ? "s" : ""}
        </div>
        <button
          onClick={() => {
            setEditingId(null);
            setShowForm((s) => !s);
          }}
          className="flex items-center gap-1.5 text-xs px-3 py-2 ml-auto"
          style={{ background: TOKENS.ink, color: TOKENS.paper, fontFamily: "'Inter', sans-serif", fontWeight: 600 }}
        >
          <Plus size={13} /> Ajouter un AO trouvé
        </button>
      </div>

      {editingAO && (
        <AddAOForm
          platforms={platforms}
          initial={editingAO}
          onCancel={() => setEditingId(null)}
          onSubmit={(data) => {
            onUpdateAO(editingAO.id, data);
            setEditingId(null);
          }}
        />
      )}

      {showForm && (
        <AddAOForm
          platforms={platforms}
          onCancel={() => setShowForm(false)}
          onSubmit={(data) => {
            onAddAO(data);
            setShowForm(false);
          }}
        />
      )}

      <div className="grid sm:grid-cols-2 gap-4">
        {filtered.map((ao) => {
          const isFollowed = followed.includes(ao.id);
          const left = daysLeft(ao.dateLimit);
          return (
            <div
              key={ao.id}
              className="p-4 border flex flex-col gap-3"
              style={{ borderColor: TOKENS.line, background: "white" }}
            >
              <div className="flex items-start justify-between gap-2">
                <div
                  className="text-[10px] uppercase tracking-wider"
                  style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}
                >
                  {ao.custom ? "Ajouté manuellement" : ao.id}
                </div>
                <div className="flex items-center gap-1.5">
                  <Tag tone={ao.type === "conception" ? "blue" : "clay"}>
                    {TYPE_LABEL[ao.type]}
                  </Tag>
                  {ao.nouveau && <Tag tone="rust">Nouveau</Tag>}
                </div>
              </div>

              <h3
                className="text-base leading-snug"
                style={{ color: TOKENS.ink, fontFamily: "'Fraunces', serif", fontWeight: 600 }}
              >
                {ao.titre}
              </h3>

              {ao.sourceName && (
                <div className="text-[11px]" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
                  Trouvé sur {ao.sourceName}
                </div>
              )}

              <div className="flex items-center gap-1.5 text-xs" style={{ color: TOKENS.inkSoft }}>
                <Building2 size={13} /> {ao.acheteur}
              </div>
              <div className="flex items-center gap-1.5 text-xs" style={{ color: TOKENS.inkSoft }}>
                <MapPin size={13} /> {ao.zone}
              </div>
              <div className="flex items-center gap-1.5 text-xs" style={{ color: TOKENS.inkSoft }}>
                <Euro size={13} /> {ao.montant ?? "Montant non communiqué"}
              </div>
              {left !== null ? (
                <div
                  className="flex items-center gap-1.5 text-xs"
                  style={{ color: left <= 20 ? TOKENS.rust : TOKENS.inkSoft }}
                >
                  <Clock size={13} /> Dépôt avant le {new Date(ao.dateLimit).toLocaleDateString("fr-FR")} · {left} j restants
                </div>
              ) : (
                <div className="flex items-center gap-1.5 text-xs" style={{ color: TOKENS.clay }}>
                  <Clock size={13} /> Date limite à confirmer sur la source
                </div>
              )}

              {ao.infoAVerifier && (
                <div
                  className="flex items-start gap-1.5 text-[11px] p-2"
                  style={{ background: TOKENS.clayDim, color: TOKENS.clay, fontFamily: "'Inter', sans-serif" }}
                >
                  <AlertTriangle size={13} className="shrink-0 mt-0.5" />
                  <span>
                    {ao.note}{" "}
                    {ao.sourceUrl && (
                      <a href={ao.sourceUrl} target="_blank" rel="noreferrer" className="underline inline-flex items-center gap-0.5">
                        Vérifier la source <ExternalLink size={10} />
                      </a>
                    )}
                  </span>
                </div>
              )}

              <div className="flex flex-wrap gap-1.5 mt-1">
                {ao.tags.map((t) => (
                  <Tag key={t} tone="moss">{t}</Tag>
                ))}
              </div>

              <div className="flex items-center gap-3 pt-1 border-t" style={{ borderColor: TOKENS.line }}>
                <button
                  onClick={() => {
                    setShowForm(false);
                    setEditingId(ao.id);
                  }}
                  className="flex items-center gap-1 text-[11px]"
                  style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}
                >
                  <Pencil size={12} /> Modifier
                </button>
                <button
                  onClick={() => onArchiveAO(ao.id)}
                  className="flex items-center gap-1 text-[11px]"
                  style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}
                >
                  <Archive size={12} /> Archiver (réalisé)
                </button>
                {ao.dateLimit && (
                  <button
                    onClick={() => sendDeadline(ao)}
                    disabled={sendingId === ao.id}
                    className="flex items-center gap-1 text-[11px]"
                    style={{ color: sentIds[ao.id] ? TOKENS.moss : TOKENS.blue, fontFamily: "'Inter', sans-serif" }}
                  >
                    {sendingId === ao.id ? (
                      <Loader2 size={12} className="animate-spin" />
                    ) : sentIds[ao.id] ? (
                      <Check size={12} />
                    ) : (
                      <CalendarDays size={12} />
                    )}
                    {sentIds[ao.id] ? "Envoyé" : "Date limite → Google Agenda"}
                  </button>
                )}
                <button
                  onClick={() => {
                    if (window.confirm("Supprimer définitivement cet AO ?")) onDeleteAO(ao.id);
                  }}
                  className="flex items-center gap-1 text-[11px] ml-auto"
                  style={{ color: TOKENS.rust, fontFamily: "'Inter', sans-serif" }}
                >
                  <Trash2 size={12} /> Supprimer
                </button>
              </div>

              <button
                onClick={() => onFollow(ao.id)}
                disabled={isFollowed}
                className="mt-1 flex items-center justify-center gap-2 text-sm py-2 transition-opacity"
                style={{
                  background: isFollowed ? TOKENS.paperDim : TOKENS.ink,
                  color: isFollowed ? TOKENS.inkSoft : TOKENS.paper,
                  fontFamily: "'Inter', sans-serif",
                  fontWeight: 600,
                }}
              >
                {isFollowed ? (
                  <>
                    <Check size={14} /> Suivi activé
                  </>
                ) : (
                  <>
                    <Plus size={14} /> Suivre cet appel d'offres
                  </>
                )}
              </button>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   REDACTION
--------------------------------------------------------- */

function Redaction({ aos, followed, projectData, ensureProject, updateProject }) {
  const followedAOs = aos.filter((a) => followed.includes(a.id));
  const [selectedId, setSelectedId] = useState(null);

  useEffect(() => {
    followedAOs.forEach((a) => ensureProject(a));
  }, [followed]);

  useEffect(() => {
    if (!selectedId && followedAOs[0]) setSelectedId(followedAOs[0].id);
  }, [followedAOs.map((a) => a.id).join(",")]);

  const selected = aos.find((a) => a.id === selectedId);
  const detail = selected ? projectData[selected.id] : null;

  const updatePieceStatus = (section, pieceId, status) =>
    updateProject(selected.id, (d) => ({
      ...d,
      [section]: d[section].map((pc) => (pc.id === pieceId ? { ...pc, status } : pc)),
    }));
  const updatePieceLabel = (section, pieceId, label) =>
    updateProject(selected.id, (d) => ({
      ...d,
      [section]: d[section].map((pc) => (pc.id === pieceId ? { ...pc, label } : pc)),
    }));
  const addPiece = (section) =>
    updateProject(selected.id, (d) => ({
      ...d,
      [section]: [...d[section], { id: "custom-" + Date.now(), label: "", status: "a_faire" }],
    }));
  const removePiece = (section, pieceId) =>
    updateProject(selected.id, (d) => ({ ...d, [section]: d[section].filter((pc) => pc.id !== pieceId) }));
  const updateMemo = (key, value) =>
    updateProject(selected.id, (d) => ({ ...d, memo: { ...d.memo, [key]: value } }));

  const inputStyle = {
    borderColor: TOKENS.line,
    fontFamily: "'Inter', sans-serif",
    color: TOKENS.ink,
    background: "white",
  };

  if (followedAOs.length === 0) {
    return (
      <div className="p-10 text-center">
        <p style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
          Suivez un appel d'offres depuis le module Veille pour préparer son dossier de réponse.
        </p>
      </div>
    );
  }

  const renderPieces = (section, list, activeColor) => (
    <div className="flex flex-col gap-1.5">
      {(list ?? []).map((piece) => {
        const isNotAsked = piece.status === "pas_demande";
        return (
          <div
            key={piece.id}
            className="flex items-center gap-2 p-1"
            style={{ opacity: isNotAsked ? 0.45 : 1, background: isNotAsked ? TOKENS.paperDim : "transparent" }}
          >
            <select
              value={piece.status ?? "a_faire"}
              onChange={(e) => updatePieceStatus(section, piece.id, e.target.value)}
              className="text-[11px] px-1.5 py-1 border outline-none shrink-0"
              style={{
                borderColor: TOKENS.line,
                fontFamily: "'JetBrains Mono', monospace",
                color: activeColor,
                background: "white",
                width: 108,
              }}
            >
              {PIECE_STATUTS.map((s) => (
                <option key={s.value} value={s.value}>{s.label}</option>
              ))}
            </select>
            <input
              value={piece.label}
              onChange={(e) => updatePieceLabel(section, piece.id, e.target.value)}
              className="flex-1 text-sm px-2 py-1 border outline-none"
              style={{
                ...inputStyle,
                textDecoration: piece.status === "termine" ? "line-through" : "none",
              }}
            />
            <button onClick={() => removePiece(section, piece.id)} style={{ color: TOKENS.inkSoft }} aria-label="Supprimer cette pièce">
              <X size={13} />
            </button>
          </div>
        );
      })}
    </div>
  );

  return (
    <div className="grid grid-cols-12">
      <div className="col-span-4 sm:col-span-3 border-r" style={{ borderColor: TOKENS.line }}>
        {followedAOs.map((a) => (
          <button
            key={a.id}
            onClick={() => setSelectedId(a.id)}
            className="w-full text-left px-4 py-3 border-b"
            style={{
              borderColor: TOKENS.line,
              background: selectedId === a.id ? TOKENS.mossDim : "transparent",
            }}
          >
            <div
              className="text-[10px] uppercase tracking-wider mb-1"
              style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}
            >
              {a.id}
            </div>
            <div
              className="text-sm leading-snug"
              style={{ color: TOKENS.ink, fontFamily: "'Inter', sans-serif", fontWeight: 500 }}
            >
              {a.titre}
            </div>
          </button>
        ))}
      </div>

      <div className="col-span-8 sm:col-span-9 p-6">
        {selected && !detail && (
          <div className="flex items-center gap-2" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
            <Loader2 size={14} className="animate-spin" /> Chargement…
          </div>
        )}
        {selected && detail && (
          <>
            <div className="flex items-center gap-2 mb-1">
              <h2
                className="text-lg"
                style={{ color: TOKENS.ink, fontFamily: "'Fraunces', serif", fontWeight: 600 }}
              >
                {selected.titre}
              </h2>
              <Tag tone={selected.type === "conception" ? "blue" : "clay"}>
                {TYPE_LABEL[selected.type]}
              </Tag>
            </div>
            <p className="text-xs mb-6" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>
              {selected.acheteur} — dépôt avant le {selected.dateLimit ? new Date(selected.dateLimit).toLocaleDateString("fr-FR") : "date à confirmer"}
            </p>

            <div className="mb-8">
              <h3
                className="text-xs uppercase tracking-wider mb-3"
                style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}
              >
                Pièces administratives
              </h3>
              {renderPieces("piecesAdmin", detail.piecesAdmin, TOKENS.moss)}
              <button
                onClick={() => addPiece("piecesAdmin")}
                className="flex items-center gap-1.5 text-[11px] px-2 py-1 mt-2"
                style={{ background: TOKENS.mossDim, color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}
              >
                <Plus size={12} /> Ajouter une pièce
              </button>
            </div>

            <div className="mb-8">
              <h3
                className="text-xs uppercase tracking-wider mb-1"
                style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}
              >
                {selected.type === "conception" ? "Pièces techniques et graphiques" : "Pièces techniques"}
              </h3>
              <p className="text-[11px] mb-3" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
                {selected.type === "conception"
                  ? "Ce dossier de conception nécessite des rendus graphiques en plus des pièces administratives."
                  : "Ce dossier de travaux/entretien ne nécessite pas de pièces graphiques."}
              </p>
              {renderPieces("piecesTech", detail.piecesTech, TOKENS.blue)}
              <button
                onClick={() => addPiece("piecesTech")}
                className="flex items-center gap-1.5 text-[11px] px-2 py-1 mt-2"
                style={{ background: TOKENS.blueDim, color: TOKENS.blue, fontFamily: "'JetBrains Mono', monospace" }}
              >
                <Plus size={12} /> Ajouter une pièce
              </button>
            </div>

            <div>
              <h3
                className="text-xs uppercase tracking-wider mb-3"
                style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}
              >
                Mémoire technique
              </h3>
              <div className="flex flex-col gap-4">
                {MEMO_SECTIONS_PAR_TYPE[selected.type].map((s) => (
                  <div key={s.key}>
                    <label
                      className="text-xs mb-1 block"
                      style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif", fontWeight: 500 }}
                    >
                      {s.label}
                    </label>
                    <textarea
                      rows={2}
                      value={detail.memo?.[s.key] ?? ""}
                      onChange={(e) => updateMemo(s.key, e.target.value)}
                      placeholder="Rédiger ou coller le contenu de cette section…"
                      className="w-full p-2 text-sm border outline-none resize-none"
                      style={inputStyle}
                    />
                  </div>
                ))}
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   SUIVI
--------------------------------------------------------- */

function Suivi({ aos, followed, statuts, onChangeStatut, onUnfollow, onOpenProject }) {
  const followedAOs = aos.filter((a) => followed.includes(a.id));

  if (followedAOs.length === 0) {
    return (
      <div className="p-10 text-center">
        <p style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
          Aucun appel d'offres suivi pour le moment.
        </p>
      </div>
    );
  }

  return (
    <div className="p-6 grid grid-cols-1 sm:grid-cols-4 gap-4">
      {STATUTS.map((col) => {
        const items = followedAOs.filter((a) => (statuts[a.id] ?? "Repéré") === col);
        return (
          <div key={col}>
            <div
              className="flex items-center justify-between px-1 py-2 border-b-2 mb-3"
              style={{ borderColor: TOKENS.ink }}
            >
              <span
                className="text-xs uppercase tracking-wider"
                style={{ color: TOKENS.ink, fontFamily: "'JetBrains Mono', monospace", fontWeight: 500 }}
              >
                {col}
              </span>
              <span
                className="text-xs"
                style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}
              >
                {items.length}
              </span>
            </div>
            <div className="flex flex-col gap-3">
              {items.map((a) => {
                const left = daysLeft(a.dateLimit);
                return (
                  <div
                    key={a.id}
                    onClick={() => onOpenProject(a)}
                    className="p-3 border relative cursor-pointer transition-shadow hover:shadow-sm"
                    style={{ borderColor: TOKENS.line, background: "white" }}
                  >
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        onUnfollow(a.id);
                      }}
                      className="absolute top-2 right-2"
                      style={{ color: TOKENS.inkSoft }}
                      aria-label="Retirer du suivi"
                    >
                      <X size={13} />
                    </button>
                    <div
                      className="text-[10px] uppercase tracking-wider mb-1"
                      style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}
                    >
                      {a.id}
                    </div>
                    <div
                      className="text-sm leading-snug mb-2 pr-4"
                      style={{ color: TOKENS.ink, fontFamily: "'Inter', sans-serif", fontWeight: 500 }}
                    >
                      {a.titre}
                    </div>
                    <div
                      className="text-[11px] mb-3 flex items-center gap-1"
                      style={{ color: left !== null && left <= 10 ? TOKENS.rust : TOKENS.inkSoft }}
                    >
                      <Calendar size={11} /> {left !== null ? `${left} j restants` : "Date à confirmer"}
                    </div>
                    <select
                      value={statuts[a.id] ?? "Repéré"}
                      onClick={(e) => e.stopPropagation()}
                      onChange={(e) => onChangeStatut(a.id, e.target.value)}
                      className="w-full text-xs px-2 py-1.5 border outline-none"
                      style={{
                        borderColor: TOKENS.line,
                        fontFamily: "'JetBrains Mono', monospace",
                        color: TOKENS.moss,
                        background: TOKENS.mossDim,
                      }}
                    >
                      {STATUTS.map((s) => (
                        <option key={s} value={s}>
                          {s}
                        </option>
                      ))}
                    </select>
                  </div>
                );
              })}
              {items.length === 0 && (
                <div
                  className="text-xs italic px-1"
                  style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}
                >
                  —
                </div>
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ---------------------------------------------------------
   SOURCES
--------------------------------------------------------- */

function Sources({ platforms, onToggle, onRemove, onAdd }) {
  const [name, setName] = useState("");
  const [url, setUrl] = useState("");

  const submit = () => {
    if (!name.trim() || !url.trim()) return;
    onAdd({ id: "custom-" + Date.now(), name: name.trim(), url: url.trim(), active: true });
    setName("");
    setUrl("");
  };

  return (
    <div className="p-6 max-w-2xl">
      <h2
        className="text-base mb-1"
        style={{ color: TOKENS.ink, fontFamily: "'Fraunces', serif", fontWeight: 600 }}
      >
        Plateformes surveillées
      </h2>
      <p className="text-xs mb-5" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
        Activez les plateformes à interroger pour la veille des appels d'offres. Dans ce prototype, l'activation
        ne déclenche pas encore de récupération réelle des annonces.
      </p>

      <div className="flex flex-col gap-2 mb-6">
        {platforms.map((p) => (
          <div
            key={p.id}
            className="flex items-center gap-3 p-3 border"
            style={{ borderColor: TOKENS.line, background: "white" }}
          >
            <button
              onClick={() => onToggle(p.id)}
              className="w-9 h-5 rounded-full relative shrink-0 transition-colors"
              style={{ background: p.active ? TOKENS.moss : TOKENS.paperDim }}
              aria-label={p.active ? "Désactiver" : "Activer"}
            >
              <span
                className="absolute top-0.5 w-4 h-4 rounded-full bg-white transition-all"
                style={{ left: p.active ? "18px" : "2px" }}
              />
            </button>
            <div className="flex-1 min-w-0">
              <div
                className="text-sm"
                style={{ color: TOKENS.ink, fontFamily: "'Inter', sans-serif", fontWeight: 500 }}
              >
                {p.name}
              </div>
              <a
                href={p.url}
                target="_blank"
                rel="noreferrer"
                className="text-[11px] inline-flex items-center gap-1"
                style={{ color: TOKENS.blue, fontFamily: "'JetBrains Mono', monospace" }}
              >
                {p.url.replace(/^https?:\/\//, "")} <ExternalLink size={10} />
              </a>
            </div>
            <button onClick={() => onRemove(p.id)} style={{ color: TOKENS.inkSoft }} aria-label="Retirer">
              <X size={15} />
            </button>
          </div>
        ))}
      </div>

      <h3
        className="text-xs uppercase tracking-wider mb-3"
        style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}
      >
        Ajouter une plateforme
      </h3>
      <div className="flex flex-col gap-2 p-3 border" style={{ borderColor: TOKENS.line, background: "white" }}>
        <input
          value={name}
          onChange={(e) => setName(e.target.value)}
          placeholder="Nom de la plateforme (ex. AWS-Achat)"
          className="w-full p-2 text-sm border outline-none"
          style={{ borderColor: TOKENS.line, fontFamily: "'Inter', sans-serif", color: TOKENS.ink }}
        />
        <input
          value={url}
          onChange={(e) => setUrl(e.target.value)}
          placeholder="URL (ex. https://www.aws-achat.info)"
          className="w-full p-2 text-sm border outline-none"
          style={{ borderColor: TOKENS.line, fontFamily: "'JetBrains Mono', monospace", color: TOKENS.ink }}
        />
        <button
          onClick={submit}
          className="flex items-center justify-center gap-2 text-sm py-2"
          style={{ background: TOKENS.ink, color: TOKENS.paper, fontFamily: "'Inter', sans-serif", fontWeight: 600 }}
        >
          <Plus size={14} /> Ajouter la plateforme
        </button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   PROJECT DETAIL (équipe, Gantt, réunions)
--------------------------------------------------------- */

function ProjectDetail({
  ao,
  detail,
  onUpdateTask,
  onAddTask,
  onRemoveTask,
  onUpdateMember,
  onAddMember,
  onRemoveMember,
  onSetMandataire,
  onUpdateLinks,
  onAddMeeting,
  onUpdateMeeting,
  onRemoveMeeting,
  onClose,
  googleConnected,
  googleToken,
  connectGoogle,
}) {
  const colorFor = (id) => {
    const idx = detail.team.findIndex((m) => m.id === id);
    return MEMBER_COLORS[idx % MEMBER_COLORS.length];
  };
  const nameFor = (id) => detail.team.find((m) => m.id === id)?.name || "Sans nom";
  const [sendingMeetingId, setSendingMeetingId] = useState(null);
  const [sentMeetingIds, setSentMeetingIds] = useState({});

  const sendMeetingToGoogle = async (r) => {
    if (!googleConnected) {
      connectGoogle();
      return;
    }
    setSendingMeetingId(r.id);
    try {
      await pushEventToGoogle(googleToken, {
        title: r.title + " — " + ao.titre,
        description: "Participants : " + ((r.attendees ?? []).map(nameFor).join(", ") || "à définir"),
        dateStr: r.date,
        time: r.time,
      });
      setSentMeetingIds((s) => ({ ...s, [r.id]: true }));
    } catch (err) {
      alert("Échec de l'envoi vers Google Agenda : " + err.message);
    } finally {
      setSendingMeetingId(null);
    }
  };

  const [viewMonth, setViewMonth] = useState(() => {
    const d = new Date(detail.rangeStart ?? "2026-08-06");
    return new Date(d.getFullYear(), d.getMonth(), 1);
  });
  const [selectedTaskId, setSelectedTaskId] = useState(null);
  const [selectedMeetingId, setSelectedMeetingId] = useState(null);

  const daysInMonth = new Date(viewMonth.getFullYear(), viewMonth.getMonth() + 1, 0).getDate();
  const monthDays = Array.from({ length: daysInMonth }, (_, i) => new Date(viewMonth.getFullYear(), viewMonth.getMonth(), i + 1));
  const monthStart = monthDays[0];
  const monthEnd = monthDays[monthDays.length - 1];
  const monthLabel = viewMonth.toLocaleDateString("fr-FR", { month: "long", year: "numeric" });
  const monthLabelCap = monthLabel.charAt(0).toUpperCase() + monthLabel.slice(1);
  const DAY_W = 34;

  const barFor = (t) => {
    const tStart = new Date(t.start);
    const tEnd = new Date(t.end);
    if (tEnd < monthStart || tStart > monthEnd) return null;
    const clipStart = tStart < monthStart ? monthStart : tStart;
    const clipEnd = tEnd > monthEnd ? monthEnd : tEnd;
    const dayIndexStart = Math.round((clipStart - monthStart) / 86400000);
    const spanDays = Math.round((clipEnd - clipStart) / 86400000) + 1;
    return { leftPct: (dayIndexStart / daysInMonth) * 100, widthPct: (spanDays / daysInMonth) * 100 };
  };
  const meetingDotFor = (r) => {
    const d = new Date(r.date);
    if (d < monthStart || d > monthEnd) return null;
    const dayIndex = Math.round((d - monthStart) / 86400000);
    return { leftPct: (dayIndex / daysInMonth) * 100 };
  };

  const selectedTask = detail.tasks.find((t) => t.id === selectedTaskId);
  const selectedMeeting = detail.meetings.find((r) => r.id === selectedMeetingId);

  const inputStyle = {
    borderColor: TOKENS.line,
    fontFamily: "'Inter', sans-serif",
    color: TOKENS.ink,
    background: "white",
  };

  return (
    <div
      className="fixed inset-0 z-50 flex items-start sm:items-center justify-center p-0 sm:p-6"
      style={{ background: "rgba(32,43,33,0.55)" }}
      onClick={onClose}
    >
      <div
        className="w-full sm:max-w-3xl max-h-full overflow-y-auto"
        style={{ background: TOKENS.paper, border: `1px solid ${TOKENS.ink}` }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-start justify-between gap-3 px-6 py-4 border-b" style={{ borderColor: TOKENS.ink }}>
          <div>
            <div
              className="text-[10px] uppercase tracking-wider mb-1"
              style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}
            >
              Fiche projet — {ao.id}
            </div>
            <h2 className="text-lg leading-snug" style={{ color: TOKENS.ink, fontFamily: "'Fraunces', serif", fontWeight: 600 }}>
              {ao.titre}
            </h2>
          </div>
          <button onClick={onClose} style={{ color: TOKENS.inkSoft }} aria-label="Fermer">
            <X size={18} />
          </button>
        </div>

        <div className="p-6">
          {/* Équipe */}
          <h3 className="text-xs uppercase tracking-wider mb-3 flex items-center gap-2" style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}>
            <Users size={13} /> Équipe
          </h3>
          <p className="text-[11px] mb-3" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
            Chaque personne renseigne elle-même son nom et son poste. Le mandataire désigne qui représente le
            groupement sur ce dossier.
          </p>
          <div className="flex flex-col gap-2 mb-4">
            {detail.team.map((m) => (
              <div key={m.id} className="flex items-center gap-2 p-2 border" style={{ borderColor: TOKENS.line, background: "white" }}>
                <span className="w-2.5 h-2.5 rounded-full shrink-0" style={{ background: colorFor(m.id) }} />
                <input
                  value={m.name}
                  onChange={(e) => onUpdateMember(m.id, { name: e.target.value })}
                  placeholder="Nom de la personne"
                  className="flex-1 min-w-0 text-sm px-2 py-1 border outline-none"
                  style={inputStyle}
                />
                <input
                  value={m.poste}
                  onChange={(e) => onUpdateMember(m.id, { poste: e.target.value })}
                  placeholder="Poste occupé"
                  className="flex-1 min-w-0 text-sm px-2 py-1 border outline-none"
                  style={inputStyle}
                />
                <label className="flex items-center gap-1 text-[11px] shrink-0" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
                  <input type="radio" name={"mandataire-" + ao.id} checked={!!m.mandataire} onChange={() => onSetMandataire(m.id)} />
                  Mandataire
                </label>
                <button onClick={() => onRemoveMember(m.id)} style={{ color: TOKENS.inkSoft }} aria-label="Retirer ce membre">
                  <X size={13} />
                </button>
              </div>
            ))}
          </div>
          <button
            onClick={onAddMember}
            className="flex items-center gap-1.5 text-[11px] px-2 py-1 mb-8"
            style={{ background: TOKENS.mossDim, color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}
          >
            <Plus size={12} /> Ajouter un membre
          </button>

          {/* Liens de travail */}
          <h3 className="text-xs uppercase tracking-wider mb-3" style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}>
            Liens de travail
          </h3>
          <div className="flex flex-col gap-2 mb-8">
            {[
              { key: "canva", label: "Canva" },
              { key: "onenote", label: "OneNote" },
            ].map((p) => (
              <div key={p.key} className="flex items-center gap-2">
                <span
                  className="text-xs w-16 shrink-0"
                  style={{ color: TOKENS.ink, fontFamily: "'Inter', sans-serif", fontWeight: 500 }}
                >
                  {p.label}
                </span>
                <input
                  value={detail.links?.[p.key] ?? ""}
                  onChange={(e) => onUpdateLinks({ [p.key]: e.target.value })}
                  placeholder={`Lien ${p.label} du projet…`}
                  className="flex-1 text-xs px-2 py-1.5 border outline-none"
                  style={{ ...inputStyle, fontFamily: "'JetBrains Mono', monospace" }}
                />
                {detail.links?.[p.key] && (
                  <a
                    href={detail.links[p.key]}
                    target="_blank"
                    rel="noreferrer"
                    className="text-[11px] px-2 py-1.5 flex items-center gap-1 shrink-0"
                    style={{ background: TOKENS.blueDim, color: TOKENS.blue, fontFamily: "'JetBrains Mono', monospace" }}
                  >
                    Ouvrir <ExternalLink size={11} />
                  </a>
                )}
              </div>
            ))}
          </div>

          {/* Gantt */}
          <div className="flex items-center justify-between mb-3">
            <h3 className="text-xs uppercase tracking-wider" style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}>
              Planning — qui fait quoi, pour quand
            </h3>
            <button
              onClick={() => {
                onAddTask();
              }}
              className="flex items-center gap-1.5 text-[11px] px-2 py-1"
              style={{ background: TOKENS.mossDim, color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}
            >
              <Plus size={12} /> Ajouter une tâche
            </button>
          </div>

          <div className="border mb-2" style={{ borderColor: TOKENS.line, background: "white" }}>
            {/* Barre de mois */}
            <div className="flex items-center justify-between px-3 py-2 border-b" style={{ borderColor: TOKENS.line }}>
              <button
                onClick={() => setViewMonth((m) => new Date(m.getFullYear(), m.getMonth() - 1, 1))}
                style={{ color: TOKENS.inkSoft }}
                aria-label="Mois précédent"
              >
                <ChevronLeft size={16} />
              </button>
              <span className="text-sm" style={{ color: TOKENS.ink, fontFamily: "'Fraunces', serif", fontWeight: 600 }}>
                {monthLabelCap}
              </span>
              <button
                onClick={() => setViewMonth((m) => new Date(m.getFullYear(), m.getMonth() + 1, 1))}
                style={{ color: TOKENS.inkSoft }}
                aria-label="Mois suivant"
              >
                <ChevronRight size={16} />
              </button>
            </div>

            <div className="overflow-x-auto">
              <div style={{ minWidth: 140 + 130 + daysInMonth * DAY_W }}>
                {/* En-tête jours */}
                <div className="grid" style={{ gridTemplateColumns: `140px 130px repeat(${daysInMonth}, ${DAY_W}px)` }}>
                  <div className="px-2 py-1.5 border-b border-r text-[10px] uppercase" style={{ borderColor: TOKENS.line, color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>
                    Personne
                  </div>
                  <div className="px-2 py-1.5 border-b border-r text-[10px] uppercase" style={{ borderColor: TOKENS.line, color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>
                    Tâches
                  </div>
                  {monthDays.map((d) => {
                    const isToday = d.toISOString().slice(0, 10) === "2026-08-06";
                    const isWeekend = d.getDay() === 0 || d.getDay() === 6;
                    return (
                      <div
                        key={d.toISOString()}
                        className="py-1.5 text-center border-b"
                        style={{
                          borderColor: TOKENS.line,
                          background: isToday ? TOKENS.clayDim : isWeekend ? TOKENS.paperDim : "transparent",
                        }}
                      >
                        <div className="text-[9px] uppercase" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>
                          {WEEKDAY_LETTERS[(d.getDay() + 6) % 7]}
                        </div>
                        <div
                          className="text-[11px]"
                          style={{ color: isToday ? TOKENS.rust : TOKENS.ink, fontFamily: "'JetBrains Mono', monospace", fontWeight: isToday ? 700 : 500 }}
                        >
                          {d.getDate()}
                        </div>
                      </div>
                    );
                  })}

                  {/* Lignes par personne */}
                  {detail.team.map((m) => {
                    const memberTasks = detail.tasks.filter((t) => (t.assigneeIds ?? []).includes(m.id));
                    const memberMeetings = detail.meetings.filter((r) => r.attendees?.includes(m.id));
                    const tasksHeight = Math.max(memberTasks.length, 1) * 26 + 12;
                    const meetingsHeight = memberMeetings.length > 0 ? 24 : 0;
                    return (
                      <React.Fragment key={m.id}>
                        <div className="px-2 py-3 border-b border-r flex items-center gap-2" style={{ borderColor: TOKENS.line }}>
                          <span className="w-2 h-2 rounded-full shrink-0" style={{ background: colorFor(m.id) }} />
                          <span className="text-xs leading-tight" style={{ color: TOKENS.ink, fontFamily: "'Inter', sans-serif", fontWeight: 500 }}>
                            {m.name || "Sans nom"}
                          </span>
                        </div>
                        <div className="px-2 py-2 border-b border-r flex flex-col gap-1">
                          {memberTasks.length === 0 && (
                            <span className="text-[10px]" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>—</span>
                          )}
                          {memberTasks.map((t) => (
                            <button
                              key={t.id}
                              onClick={() => setSelectedTaskId(t.id)}
                              className="text-left text-[10px] px-1.5 py-0.5 truncate"
                              style={{
                                background: selectedTaskId === t.id ? TOKENS.mossDim : "transparent",
                                color: TOKENS.ink,
                                fontFamily: "'Inter', sans-serif",
                              }}
                              title={t.label}
                            >
                              {t.label}
                              {(t.assigneeIds ?? []).length > 1 && (
                                <span style={{ color: TOKENS.inkSoft }}> · +{t.assigneeIds.length - 1}</span>
                              )}
                            </button>
                          ))}
                        </div>
                        <div
                          className="relative border-b"
                          style={{ borderColor: TOKENS.line, gridColumn: `3 / span ${daysInMonth}`, height: tasksHeight + meetingsHeight }}
                        >
                          {memberTasks.map((t, idx) => {
                            const bar = barFor(t);
                            if (!bar) return null;
                            return (
                              <button
                                key={t.id}
                                onClick={() => setSelectedTaskId(t.id)}
                                className="absolute text-[10px] px-2 flex items-center truncate"
                                style={{
                                  left: bar.leftPct + "%",
                                  width: `calc(${bar.widthPct}% - 3px)`,
                                  top: idx * 26 + 6,
                                  height: 20,
                                  background: colorFor(m.id),
                                  color: "white",
                                  fontFamily: "'Inter', sans-serif",
                                  fontWeight: 500,
                                  outline: selectedTaskId === t.id ? `2px solid ${TOKENS.ink}` : "none",
                                }}
                              >
                                {t.label}
                              </button>
                            );
                          })}
                          {memberMeetings.map((r) => {
                            const dot = meetingDotFor(r);
                            if (!dot) return null;
                            return (
                              <button
                                key={r.id}
                                onClick={() => setSelectedMeetingId(r.id)}
                                className="absolute flex items-center justify-center rounded-full"
                                title={r.title + " — " + r.time}
                                style={{
                                  left: `calc(${dot.leftPct}% + 2px)`,
                                  top: tasksHeight + 2,
                                  width: 16,
                                  height: 16,
                                  background: TOKENS.rust,
                                  outline: selectedMeetingId === r.id ? `2px solid ${TOKENS.ink}` : "none",
                                }}
                              >
                                <CalendarDays size={10} color="white" />
                              </button>
                            );
                          })}
                        </div>
                      </React.Fragment>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>

          {/* Légende */}
          <div className="flex flex-wrap items-center gap-3 mb-6">
            {detail.team.map((m) => (
              <span key={m.id} className="flex items-center gap-1.5 text-[11px]" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
                <span className="w-2 h-2 rounded-full" style={{ background: colorFor(m.id) }} /> {m.name || "Sans nom"}
              </span>
            ))}
          </div>

          {/* Panneau d'édition de la tâche sélectionnée */}
          {selectedTask && (
            <div className="p-3 border mb-8" style={{ borderColor: TOKENS.moss, background: TOKENS.mossDim }}>
              <div className="flex items-center justify-between mb-2">
                <span className="text-[10px] uppercase tracking-wider" style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}>
                  Modifier la tâche
                </span>
                <button onClick={() => setSelectedTaskId(null)} style={{ color: TOKENS.moss }} aria-label="Fermer l'édition">
                  <X size={14} />
                </button>
              </div>
              <div className="grid grid-cols-12 gap-1.5 mb-2">
                <input
                  value={selectedTask.label}
                  onChange={(e) => onUpdateTask(selectedTask.id, { label: e.target.value })}
                  className="col-span-12 sm:col-span-7 text-xs px-2 py-1 border outline-none"
                  style={inputStyle}
                />
                <input
                  type="date"
                  value={selectedTask.start}
                  onChange={(e) => onUpdateTask(selectedTask.id, { start: e.target.value })}
                  className="col-span-5 sm:col-span-2 text-xs px-1 py-1 border outline-none"
                  style={inputStyle}
                />
                <input
                  type="date"
                  value={selectedTask.end}
                  onChange={(e) => onUpdateTask(selectedTask.id, { end: e.target.value })}
                  className="col-span-6 sm:col-span-2 text-xs px-1 py-1 border outline-none"
                  style={inputStyle}
                />
                <button
                  onClick={() => {
                    onRemoveTask(selectedTask.id);
                    setSelectedTaskId(null);
                  }}
                  className="col-span-1 flex items-center justify-center"
                  style={{ color: TOKENS.rust }}
                  aria-label="Supprimer la tâche"
                >
                  <X size={13} />
                </button>
              </div>
              <div className="text-[10px] uppercase tracking-wider mb-1" style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}>
                Assignée à
              </div>
              <div className="flex flex-wrap gap-2">
                {detail.team.map((m) => {
                  const checked = (selectedTask.assigneeIds ?? []).includes(m.id);
                  return (
                    <label
                      key={m.id}
                      className="flex items-center gap-1.5 text-[11px] px-2 py-1 border cursor-pointer"
                      style={{ borderColor: TOKENS.line, background: checked ? "white" : "transparent", fontFamily: "'Inter', sans-serif" }}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={(e) => {
                          const current = selectedTask.assigneeIds ?? [];
                          const next = e.target.checked
                            ? [...current, m.id]
                            : current.filter((id) => id !== m.id);
                          onUpdateTask(selectedTask.id, { assigneeIds: next });
                        }}
                      />
                      {m.name || "Sans nom"}
                    </label>
                  );
                })}
              </div>
            </div>
          )}

          {/* Réunions */}
          <div className="flex items-center justify-between mb-3">
            <h3 className="text-xs uppercase tracking-wider flex items-center gap-2" style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}>
              <CalendarDays size={13} /> Réunions
            </h3>
            <div className="flex items-center gap-2">
              <button
                onClick={() => {
                  const id = onAddMeeting();
                  setSelectedMeetingId(id);
                }}
                className="flex items-center gap-1.5 text-[11px] px-2 py-1"
                style={{ background: TOKENS.mossDim, color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}
              >
                <Plus size={12} /> Ajouter une réunion
              </button>
              <button
                onClick={() => downloadICS(ao.id + "-reunions.ics", buildICS(detail.meetings))}
                className="flex items-center gap-1.5 text-[11px] px-2 py-1"
                style={{ background: TOKENS.blueDim, color: TOKENS.blue, fontFamily: "'JetBrains Mono', monospace" }}
              >
                <Download size={12} /> Tout exporter (.ics)
              </button>
            </div>
          </div>
          <div className="flex flex-col gap-2 mb-2">
            {detail.meetings.length === 0 && (
              <div className="text-[11px]" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>Aucune réunion planifiée.</div>
            )}
            {detail.meetings.map((r) => (
              <div
                key={r.id}
                onClick={() => setSelectedMeetingId(r.id)}
                className="flex items-center justify-between gap-3 p-3 border cursor-pointer"
                style={{ borderColor: selectedMeetingId === r.id ? TOKENS.moss : TOKENS.line, background: "white" }}
              >
                <div>
                  <div className="text-sm" style={{ color: TOKENS.ink, fontFamily: "'Inter', sans-serif", fontWeight: 500 }}>{r.title}</div>
                  <div className="text-[11px]" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>
                    {new Date(r.date).toLocaleDateString("fr-FR")} · {r.time} — {(r.attendees ?? []).map(nameFor).join(", ") || "aucun participant"}
                  </div>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      sendMeetingToGoogle(r);
                    }}
                    disabled={sendingMeetingId === r.id}
                    style={{ color: sentMeetingIds[r.id] ? TOKENS.moss : TOKENS.blue }}
                    aria-label="Envoyer vers Google Agenda"
                    title="Envoyer vers mon Google Agenda"
                  >
                    {sendingMeetingId === r.id ? (
                      <Loader2 size={15} className="animate-spin" />
                    ) : sentMeetingIds[r.id] ? (
                      <Check size={15} />
                    ) : (
                      <CalendarDays size={15} />
                    )}
                  </button>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      downloadICS(r.id + ".ics", buildICS([r]));
                    }}
                    style={{ color: TOKENS.blue }}
                    aria-label="Exporter cette réunion"
                  >
                    <Download size={15} />
                  </button>
                </div>
              </div>
            ))}
          </div>

          {selectedMeeting && (
            <div className="p-3 border mb-2" style={{ borderColor: TOKENS.rust, background: TOKENS.rustDim }}>
              <div className="flex items-center justify-between mb-2">
                <span className="text-[10px] uppercase tracking-wider" style={{ color: TOKENS.rust, fontFamily: "'JetBrains Mono', monospace" }}>
                  Modifier la réunion
                </span>
                <button onClick={() => setSelectedMeetingId(null)} style={{ color: TOKENS.rust }} aria-label="Fermer l'édition">
                  <X size={14} />
                </button>
              </div>
              <div className="grid grid-cols-12 gap-1.5 mb-2">
                <input
                  value={selectedMeeting.title}
                  onChange={(e) => onUpdateMeeting(selectedMeeting.id, { title: e.target.value })}
                  className="col-span-12 sm:col-span-6 text-xs px-2 py-1 border outline-none"
                  style={inputStyle}
                />
                <input
                  type="date"
                  value={selectedMeeting.date}
                  onChange={(e) => onUpdateMeeting(selectedMeeting.id, { date: e.target.value })}
                  className="col-span-6 sm:col-span-3 text-xs px-1 py-1 border outline-none"
                  style={inputStyle}
                />
                <input
                  type="time"
                  value={selectedMeeting.time}
                  onChange={(e) => onUpdateMeeting(selectedMeeting.id, { time: e.target.value })}
                  className="col-span-5 sm:col-span-2 text-xs px-1 py-1 border outline-none"
                  style={inputStyle}
                />
                <button
                  onClick={() => {
                    onRemoveMeeting(selectedMeeting.id);
                    setSelectedMeetingId(null);
                  }}
                  className="col-span-1 flex items-center justify-center"
                  style={{ color: TOKENS.rust }}
                  aria-label="Supprimer la réunion"
                >
                  <X size={13} />
                </button>
              </div>
              <div className="flex flex-wrap gap-2">
                {detail.team.map((m) => {
                  const checked = (selectedMeeting.attendees ?? []).includes(m.id);
                  return (
                    <label
                      key={m.id}
                      className="flex items-center gap-1.5 text-[11px] px-2 py-1 border cursor-pointer"
                      style={{ borderColor: TOKENS.line, background: checked ? TOKENS.mossDim : "white", fontFamily: "'Inter', sans-serif" }}
                    >
                      <input
                        type="checkbox"
                        checked={checked}
                        onChange={(e) => {
                          const next = e.target.checked
                            ? [...(selectedMeeting.attendees ?? []), m.id]
                            : (selectedMeeting.attendees ?? []).filter((id) => id !== m.id);
                          onUpdateMeeting(selectedMeeting.id, { attendees: next });
                        }}
                      />
                      {m.name || "Sans nom"}
                    </label>
                  );
                })}
              </div>
            </div>
          )}

          <p className="text-[11px] mt-2" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
            Le fichier .ics s'importe en un clic dans Google Agenda (ou tout autre calendrier) pour chaque membre de
            l'équipe. Une synchronisation automatique bidirectionnelle nécessiterait une connexion OAuth à Google
            Agenda côté serveur — hors périmètre de ce prototype.
          </p>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   FINANCES
--------------------------------------------------------- */

function fmtMonthKey(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function exportFinancesToExcel(ready, projectData) {
  const caRows = [];
  const hoursRows = [];

  ready.forEach((ao) => {
    const d = projectData[ao.id];
    Object.entries(d.finance?.caByMonth ?? {}).forEach(([month, val]) => {
      const n = parseFloat((val ?? "").toString().replace(/[^\d.]/g, ""));
      if (val !== "" && val !== undefined && !isNaN(n)) {
        caRows.push({ Mois: month, "AO": ao.id, Projet: ao.titre, "CA (€)": n });
      }
    });
    d.team.forEach((m) => {
      Object.entries(d.finance?.hours?.[m.id] ?? {}).forEach(([month, h]) => {
        const n = parseFloat(h);
        if (h !== "" && h !== undefined && !isNaN(n) && n !== 0) {
          hoursRows.push({
            Mois: month,
            Personne: m.name || "Sans nom",
            Poste: m.poste || "",
            "AO": ao.id,
            Projet: ao.titre,
            Heures: n,
          });
        }
      });
    });
  });

  const months = Array.from(new Set([...caRows.map((r) => r.Mois), ...hoursRows.map((r) => r.Mois)])).sort();
  const synthese = months.map((month) => ({
    Mois: month,
    "CA total (€)": caRows.filter((r) => r.Mois === month).reduce((s, r) => s + r["CA (€)"], 0),
    "Heures totales": hoursRows.filter((r) => r.Mois === month).reduce((s, r) => s + r.Heures, 0),
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(synthese), "Synthèse par mois");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(caRows), "CA par projet");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(hoursRows), "Heures par personne");
  XLSX.writeFile(wb, "suivi-financier-ao-paysage.xlsx");
}

function Finances({ followedAOs, projectData, ensureProject, onUpdateCA, onUpdateHours }) {
  useEffect(() => {
    followedAOs.forEach((ao) => ensureProject(ao));
  }, [followedAOs]);

  const [viewMonth, setViewMonth] = useState(new Date(2026, 7, 1)); // Août 2026
  const monthKey = fmtMonthKey(viewMonth);
  const monthLabel = viewMonth.toLocaleDateString("fr-FR", { month: "long", year: "numeric" });
  const monthLabelCap = monthLabel.charAt(0).toUpperCase() + monthLabel.slice(1);

  if (followedAOs.length === 0) {
    return (
      <div className="p-10 text-center">
        <p style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
          Suivez un appel d'offres pour commencer à saisir son CA et les heures passées par l'équipe.
        </p>
      </div>
    );
  }

  const ready = followedAOs.filter((ao) => projectData[ao.id]);

  const sumAllMonths = (obj) =>
    Object.values(obj ?? {}).reduce((s, v) => s + (parseFloat((v ?? "").toString().replace(/[^\d.]/g, "")) || 0), 0);

  const caDuMois = ready.reduce((sum, ao) => {
    const v = parseFloat((projectData[ao.id].finance?.caByMonth?.[monthKey] ?? "").toString().replace(/[^\d.]/g, ""));
    return sum + (isNaN(v) ? 0 : v);
  }, 0);
  const caCumule = ready.reduce((sum, ao) => sum + sumAllMonths(projectData[ao.id].finance?.caByMonth), 0);

  // Agrégat des heures par nom pour le mois affiché, tous projets suivis confondus
  const hoursByName = {};
  const hoursByNameCumule = {};
  ready.forEach((ao) => {
    const d = projectData[ao.id];
    d.team.forEach((m) => {
      const key = (m.name || "Sans nom").trim();
      const hMonth = parseFloat(d.finance?.hours?.[m.id]?.[monthKey] ?? 0) || 0;
      hoursByName[key] = (hoursByName[key] ?? 0) + hMonth;
      hoursByNameCumule[key] = (hoursByNameCumule[key] ?? 0) + sumAllMonths(d.finance?.hours?.[m.id]);
    });
  });
  const hoursSorted = Object.entries(hoursByName).sort((a, b) => b[1] - a[1]);

  return (
    <div className="p-6">
      <div className="flex items-center justify-between gap-3 mb-5">
        <div className="flex items-center justify-between flex-1 px-3 py-2 border" style={{ borderColor: TOKENS.line, background: "white" }}>
          <button onClick={() => setViewMonth((m) => new Date(m.getFullYear(), m.getMonth() - 1, 1))} style={{ color: TOKENS.inkSoft }} aria-label="Mois précédent">
            <ChevronLeft size={16} />
          </button>
          <span className="text-sm" style={{ color: TOKENS.ink, fontFamily: "'Fraunces', serif", fontWeight: 600 }}>
            {monthLabelCap}
          </span>
          <button onClick={() => setViewMonth((m) => new Date(m.getFullYear(), m.getMonth() + 1, 1))} style={{ color: TOKENS.inkSoft }} aria-label="Mois suivant">
            <ChevronRight size={16} />
          </button>
        </div>
        <button
          onClick={() => exportFinancesToExcel(ready, projectData)}
          className="flex items-center gap-1.5 text-xs px-3 py-2 shrink-0"
          style={{ background: TOKENS.ink, color: TOKENS.paper, fontFamily: "'Inter', sans-serif", fontWeight: 600 }}
        >
          <FileSpreadsheet size={14} /> Exporter en Excel
        </button>
      </div>

      <div className="grid sm:grid-cols-2 gap-4 mb-8">
        <div className="p-4 border" style={{ borderColor: TOKENS.line, background: "white" }}>
          <div className="text-[10px] uppercase tracking-wider mb-1" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>
            CA du mois — {monthLabelCap}
          </div>
          <div className="text-2xl" style={{ color: TOKENS.ink, fontFamily: "'Fraunces', serif", fontWeight: 600 }}>
            {caDuMois.toLocaleString("fr-FR")} €
          </div>
          <div className="text-[11px] mt-1" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
            Cumulé toutes périodes : {caCumule.toLocaleString("fr-FR")} €
          </div>
        </div>
        <div className="p-4 border" style={{ borderColor: TOKENS.line, background: "white" }}>
          <div className="text-[10px] uppercase tracking-wider mb-2" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>
            Heures par personne — {monthLabelCap}
          </div>
          {hoursSorted.length === 0 ? (
            <div className="text-xs" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>—</div>
          ) : (
            <div className="flex flex-col gap-1">
              {hoursSorted.map(([name, h]) => (
                <div key={name} className="flex items-center justify-between text-sm">
                  <span style={{ color: TOKENS.ink, fontFamily: "'Inter', sans-serif", fontWeight: 500 }}>{name}</span>
                  <span style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}>
                    {h} h
                    <span style={{ color: TOKENS.inkSoft }}> · {hoursByNameCumule[name]} h cumulé</span>
                  </span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      <div className="flex flex-col gap-4">
        {ready.map((ao) => {
          const d = projectData[ao.id];
          const projectHoursMonth = d.team.reduce(
            (s, m) => s + (parseFloat(d.finance?.hours?.[m.id]?.[monthKey] ?? 0) || 0),
            0
          );
          return (
            <div key={ao.id} className="border" style={{ borderColor: TOKENS.line, background: "white" }}>
              <div className="flex items-start justify-between gap-3 px-4 py-3 border-b" style={{ borderColor: TOKENS.line }}>
                <div>
                  <div className="text-[10px] uppercase tracking-wider mb-1" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>
                    {ao.id}
                  </div>
                  <div className="text-sm" style={{ color: TOKENS.ink, fontFamily: "'Inter', sans-serif", fontWeight: 500 }}>
                    {ao.titre}
                  </div>
                </div>
                <div className="flex items-center gap-2 shrink-0">
                  <span className="text-xs" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>CA {monthLabelCap}</span>
                  <input
                    value={d.finance?.caByMonth?.[monthKey] ?? ""}
                    onChange={(e) => onUpdateCA(ao.id, monthKey, e.target.value)}
                    placeholder="0"
                    className="w-24 text-sm px-2 py-1 border outline-none text-right"
                    style={{ borderColor: TOKENS.line, fontFamily: "'JetBrains Mono', monospace", color: TOKENS.ink }}
                  />
                  <span className="text-xs" style={{ color: TOKENS.inkSoft }}>€</span>
                </div>
              </div>
              <div className="px-4 py-3">
                <div className="text-[10px] uppercase tracking-wider mb-2" style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}>
                  Heures effectuées en {monthLabelCap} (saisie manuelle par personne)
                </div>
                <div className="flex flex-col gap-2">
                  {d.team.map((m) => (
                    <div key={m.id} className="flex items-center gap-3">
                      <span className="flex-1 text-sm truncate" style={{ color: TOKENS.ink, fontFamily: "'Inter', sans-serif" }}>
                        {m.name || "Sans nom"}
                        {m.poste && <span style={{ color: TOKENS.inkSoft }}> — {m.poste}</span>}
                      </span>
                      <input
                        type="number"
                        min="0"
                        step="0.5"
                        value={d.finance?.hours?.[m.id]?.[monthKey] ?? ""}
                        onChange={(e) => onUpdateHours(ao.id, m.id, monthKey, e.target.value)}
                        placeholder="0"
                        className="w-20 text-sm px-2 py-1 border outline-none text-right"
                        style={{ borderColor: TOKENS.line, fontFamily: "'JetBrains Mono', monospace", color: TOKENS.ink }}
                      />
                      <span className="text-xs w-4" style={{ color: TOKENS.inkSoft }}>h</span>
                    </div>
                  ))}
                </div>
                <div className="flex items-center justify-between mt-2 pt-2 border-t" style={{ borderColor: TOKENS.line }}>
                  <span className="text-xs" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>Total heures — {monthLabelCap}</span>
                  <span className="text-xs" style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace", fontWeight: 600 }}>{projectHoursMonth} h</span>
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------
   HISTORIQUE (AO réalisés + galerie photo par dossier)
--------------------------------------------------------- */

function PhotoGallery({ ao }) {
  const [photos, setPhotos] = useState([]);
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState(false);
  const folderRef = storageRef(storage, `organisations/${ORG_ID}/photos/${ao.id}`);

  const refresh = () => {
    setLoading(true);
    listAll(folderRef)
      .then(async (res) => {
        const items = await Promise.all(
          res.items.map(async (item) => ({ name: item.name, url: await getDownloadURL(item), ref: item }))
        );
        setPhotos(items);
      })
      .catch((err) => console.error("Erreur de chargement des photos :", err))
      .finally(() => setLoading(false));
  };

  useEffect(refresh, [ao.id]);

  const handleUpload = async (e) => {
    const files = Array.from(e.target.files ?? []);
    if (files.length === 0) return;
    setUploading(true);
    try {
      for (const file of files) {
        const fileRef = storageRef(folderRef, `${Date.now()}-${file.name}`);
        await uploadBytes(fileRef, file);
      }
      refresh();
    } catch (err) {
      console.error("Erreur d'envoi de photo :", err);
    } finally {
      setUploading(false);
      e.target.value = "";
    }
  };

  const handleDelete = (photo) => {
    deleteObject(photo.ref)
      .then(() => setPhotos((p) => p.filter((x) => x.name !== photo.name)))
      .catch((err) => console.error("Erreur de suppression de photo :", err));
  };

  return (
    <div className="mt-3 pt-3 border-t" style={{ borderColor: TOKENS.line }}>
      <div className="flex items-center justify-between mb-2">
        <span className="text-[11px] uppercase tracking-wider flex items-center gap-1.5" style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}>
          <Image size={13} /> Photos du dossier
        </span>
        <label
          className="flex items-center gap-1.5 text-[11px] px-2 py-1 cursor-pointer"
          style={{ background: TOKENS.mossDim, color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}
        >
          {uploading ? <Loader2 size={12} className="animate-spin" /> : <Upload size={12} />}
          Ajouter des photos
          <input type="file" accept="image/*" multiple onChange={handleUpload} className="hidden" disabled={uploading} />
        </label>
      </div>

      {loading ? (
        <div className="text-[11px]" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>Chargement…</div>
      ) : photos.length === 0 ? (
        <div className="text-[11px]" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>Aucune photo pour l'instant.</div>
      ) : (
        <div className="grid grid-cols-3 sm:grid-cols-4 gap-2">
          {photos.map((photo) => (
            <div key={photo.name} className="relative group">
              <a href={photo.url} target="_blank" rel="noreferrer">
                <img src={photo.url} alt="" className="w-full h-20 object-cover border" style={{ borderColor: TOKENS.line }} />
              </a>
              <button
                onClick={() => handleDelete(photo)}
                className="absolute top-1 right-1 p-0.5"
                style={{ background: "rgba(32,43,33,0.7)", color: "white" }}
                aria-label="Supprimer la photo"
              >
                <X size={11} />
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function Historique({ aos, onRestoreAO }) {
  const [openId, setOpenId] = useState(null);

  if (aos.length === 0) {
    return (
      <div className="p-10 text-center">
        <p style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
          Aucun dossier archivé pour l'instant. Une fois un projet réalisé, archivez-le depuis Veille
          pour le retrouver ici avec ses photos.
        </p>
      </div>
    );
  }

  return (
    <div className="p-6 flex flex-col gap-3">
      {aos.map((ao) => (
        <div key={ao.id} className="border p-4" style={{ borderColor: TOKENS.line, background: "white" }}>
          <div className="flex items-start justify-between gap-3">
            <div className="cursor-pointer" onClick={() => setOpenId((id) => (id === ao.id ? null : ao.id))}>
              <div className="text-[10px] uppercase tracking-wider mb-1" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>
                {ao.acheteur}
              </div>
              <div className="text-sm" style={{ color: TOKENS.ink, fontFamily: "'Fraunces', serif", fontWeight: 600 }}>
                {ao.titre}
              </div>
              <div className="text-[11px] mt-0.5" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
                {ao.zone}
              </div>
            </div>
            <button
              onClick={() => onRestoreAO(ao.id)}
              className="text-[11px] px-2 py-1 shrink-0"
              style={{ background: TOKENS.paperDim, color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}
            >
              Réactiver
            </button>
          </div>
          {openId === ao.id && <PhotoGallery ao={ao} />}
        </div>
      ))}
    </div>
  );
}

/* ---------------------------------------------------------
   VUE D'ENSEMBLE (Gantt inter-AO + suivi financier par date de fin)
--------------------------------------------------------- */

const RESULTAT_LABEL = { en_attente: "En attente", gagne: "Gagné", perdu: "Perdu" };
const RESULTAT_COLOR = { en_attente: TOKENS.blue, gagne: TOKENS.moss, perdu: TOKENS.rust };

function VueEnsemble({ aos, followed, projectData, ensureProject }) {
  const followedAOs = aos.filter((a) => followed.includes(a.id));
  useEffect(() => {
    followedAOs.forEach((a) => ensureProject(a));
  }, [followed]);

  const [viewMonth, setViewMonth] = useState(() => {
    const d = new Date();
    return new Date(d.getFullYear(), d.getMonth(), 1);
  });
  const daysInMonth = new Date(viewMonth.getFullYear(), viewMonth.getMonth() + 1, 0).getDate();
  const monthDays = Array.from({ length: daysInMonth }, (_, i) => new Date(viewMonth.getFullYear(), viewMonth.getMonth(), i + 1));
  const monthStart = monthDays[0];
  const monthEnd = monthDays[monthDays.length - 1];
  const monthLabel = viewMonth.toLocaleDateString("fr-FR", { month: "long", year: "numeric" });
  const monthLabelCap = monthLabel.charAt(0).toUpperCase() + monthLabel.slice(1);
  const DAY_W = 30;
  const todayStr = new Date().toISOString().slice(0, 10);

  const planned = followedAOs.filter((a) => a.dateDebut && a.dateFin);
  const missingDates = followedAOs.filter((a) => !(a.dateDebut && a.dateFin));

  const barFor = (a) => {
    const s = new Date(a.dateDebut);
    const e = new Date(a.dateFin);
    if (e < monthStart || s > monthEnd) return null;
    const clipStart = s < monthStart ? monthStart : s;
    const clipEnd = e > monthEnd ? monthEnd : e;
    const dayIndexStart = Math.round((clipStart - monthStart) / 86400000);
    const spanDays = Math.round((clipEnd - clipStart) / 86400000) + 1;
    return { leftPct: (dayIndexStart / daysInMonth) * 100, widthPct: (spanDays / daysInMonth) * 100 };
  };
  const resultDotFor = (a) => {
    if (!a.dateResultat) return null;
    const d = new Date(a.dateResultat);
    if (d < monthStart || d > monthEnd) return null;
    const dayIndex = Math.round((d - monthStart) / 86400000);
    return { leftPct: (dayIndex / daysInMonth) * 100 };
  };

  // Suivi financier : chaque AO est rattaché au mois de sa date de fin.
  const finiCeMois = followedAOs.filter((a) => a.dateFin && a.dateFin.slice(0, 7) === fmtMonthKey(viewMonth));
  const caOf = (a) => {
    const d = projectData[a.id];
    if (!d?.finance?.caByMonth) return 0;
    return Object.values(d.finance.caByMonth).reduce((s, v) => s + (parseFloat((v ?? "").toString().replace(/[^\d.]/g, "")) || 0), 0);
  };
  const caGagne = finiCeMois.filter((a) => a.resultat === "gagne").reduce((s, a) => s + caOf(a), 0);
  const caPerdu = finiCeMois.filter((a) => a.resultat === "perdu").reduce((s, a) => s + caOf(a), 0);

  return (
    <div className="p-6">
      <div className="flex items-center justify-between mb-5 px-3 py-2 border" style={{ borderColor: TOKENS.line, background: "white" }}>
        <button onClick={() => setViewMonth((m) => new Date(m.getFullYear(), m.getMonth() - 1, 1))} style={{ color: TOKENS.inkSoft }} aria-label="Mois précédent">
          <ChevronLeft size={16} />
        </button>
        <span className="text-sm" style={{ color: TOKENS.ink, fontFamily: "'Fraunces', serif", fontWeight: 600 }}>
          {monthLabelCap}
        </span>
        <button onClick={() => setViewMonth((m) => new Date(m.getFullYear(), m.getMonth() + 1, 1))} style={{ color: TOKENS.inkSoft }} aria-label="Mois suivant">
          <ChevronRight size={16} />
        </button>
      </div>

      {/* Gantt inter-AO */}
      <h3 className="text-xs uppercase tracking-wider mb-2" style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}>
        Planning des appels d'offres
      </h3>
      {planned.length === 0 ? (
        <p className="text-xs mb-6" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
          Aucun AO suivi n'a encore de date de début/fin renseignée. Ajoutez-les depuis Veille (bouton Modifier).
        </p>
      ) : (
        <div className="border mb-2" style={{ borderColor: TOKENS.line, background: "white" }}>
          <div className="overflow-x-auto">
            <div style={{ minWidth: 180 + daysInMonth * DAY_W }}>
              <div className="grid" style={{ gridTemplateColumns: `180px repeat(${daysInMonth}, ${DAY_W}px)` }}>
                <div className="px-2 py-1.5 border-b border-r text-[10px] uppercase" style={{ borderColor: TOKENS.line, color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>
                  Appel d'offres
                </div>
                {monthDays.map((d) => {
                  const isToday = d.toISOString().slice(0, 10) === todayStr;
                  const isWeekend = d.getDay() === 0 || d.getDay() === 6;
                  return (
                    <div key={d.toISOString()} className="py-1.5 text-center border-b" style={{ borderColor: TOKENS.line, background: isToday ? TOKENS.clayDim : isWeekend ? TOKENS.paperDim : "transparent" }}>
                      <div className="text-[9px] uppercase" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>
                        {WEEKDAY_LETTERS[(d.getDay() + 6) % 7]}
                      </div>
                      <div className="text-[11px]" style={{ color: isToday ? TOKENS.rust : TOKENS.ink, fontFamily: "'JetBrains Mono', monospace", fontWeight: isToday ? 700 : 500 }}>
                        {d.getDate()}
                      </div>
                    </div>
                  );
                })}

                {planned.map((a) => {
                  const bar = barFor(a);
                  const dot = resultDotFor(a);
                  const color = RESULTAT_COLOR[a.resultat] ?? TOKENS.blue;
                  return (
                    <React.Fragment key={a.id}>
                      <div className="px-2 py-3 border-b border-r" style={{ borderColor: TOKENS.line }}>
                        <div className="text-xs leading-tight truncate" style={{ color: TOKENS.ink, fontFamily: "'Inter', sans-serif", fontWeight: 500 }} title={a.titre}>
                          {a.titre}
                        </div>
                        <Tag tone={a.resultat === "gagne" ? "moss" : a.resultat === "perdu" ? "rust" : "blue"}>
                          {RESULTAT_LABEL[a.resultat ?? "en_attente"]}
                        </Tag>
                      </div>
                      <div className="relative border-b" style={{ borderColor: TOKENS.line, gridColumn: `2 / span ${daysInMonth}`, height: 40 }}>
                        {bar && (
                          <div
                            className="absolute text-[10px] px-2 flex items-center truncate text-white"
                            style={{ left: bar.leftPct + "%", width: `calc(${bar.widthPct}% - 3px)`, top: 10, height: 20, background: color, fontFamily: "'Inter', sans-serif", fontWeight: 500 }}
                          >
                            {a.titre}
                          </div>
                        )}
                        {dot && (
                          <div
                            className="absolute flex items-center justify-center rounded-full"
                            title={"Résultat : " + new Date(a.dateResultat).toLocaleDateString("fr-FR")}
                            style={{ left: `calc(${dot.leftPct}% + 2px)`, top: 34, width: 14, height: 14, background: TOKENS.ink }}
                          >
                            <Trophy size={9} color="white" />
                          </div>
                        )}
                      </div>
                    </React.Fragment>
                  );
                })}
              </div>
            </div>
          </div>
        </div>
      )}
      <div className="flex items-center gap-4 mb-8 text-[11px]" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
        <span className="flex items-center gap-1.5"><span className="w-2 h-2 rounded-full" style={{ background: TOKENS.blue }} /> En attente</span>
        <span className="flex items-center gap-1.5"><span className="w-2 h-2 rounded-full" style={{ background: TOKENS.moss }} /> Gagné</span>
        <span className="flex items-center gap-1.5"><span className="w-2 h-2 rounded-full" style={{ background: TOKENS.rust }} /> Perdu</span>
        <span className="flex items-center gap-1.5"><Trophy size={11} /> Date de résultat</span>
      </div>

      {/* Suivi financier par date de fin */}
      <h3 className="text-xs uppercase tracking-wider mb-2" style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}>
        Suivi financier — AO se terminant en {monthLabelCap}
      </h3>
      <div className="grid sm:grid-cols-2 gap-4 mb-4">
        <div className="p-4 border" style={{ borderColor: TOKENS.moss, background: TOKENS.mossDim }}>
          <div className="text-[10px] uppercase tracking-wider mb-1" style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace" }}>CA gagné</div>
          <div className="text-2xl" style={{ color: TOKENS.ink, fontFamily: "'Fraunces', serif", fontWeight: 600 }}>{caGagne.toLocaleString("fr-FR")} €</div>
        </div>
        <div className="p-4 border" style={{ borderColor: TOKENS.rust, background: TOKENS.rustDim }}>
          <div className="text-[10px] uppercase tracking-wider mb-1" style={{ color: TOKENS.rust, fontFamily: "'JetBrains Mono', monospace" }}>CA perdu</div>
          <div className="text-2xl" style={{ color: TOKENS.ink, fontFamily: "'Fraunces', serif", fontWeight: 600 }}>{caPerdu.toLocaleString("fr-FR")} €</div>
        </div>
      </div>

      {finiCeMois.length === 0 ? (
        <p className="text-xs" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>Aucun AO suivi ne se termine ce mois-ci.</p>
      ) : (
        <div className="flex flex-col gap-2">
          {finiCeMois.map((a) => (
            <div key={a.id} className="flex items-center justify-between gap-3 p-3 border" style={{ borderColor: TOKENS.line, background: "white" }}>
              <div>
                <div className="text-sm" style={{ color: TOKENS.ink, fontFamily: "'Inter', sans-serif", fontWeight: 500 }}>{a.titre}</div>
                <div className="text-[11px]" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>
                  Fin le {new Date(a.dateFin).toLocaleDateString("fr-FR")} · CA {caOf(a).toLocaleString("fr-FR")} €
                </div>
              </div>
              <Tag tone={a.resultat === "gagne" ? "moss" : a.resultat === "perdu" ? "rust" : "blue"}>
                {RESULTAT_LABEL[a.resultat ?? "en_attente"]}
              </Tag>
            </div>
          ))}
        </div>
      )}

      {missingDates.length > 0 && (
        <p className="text-[11px] mt-6" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
          {missingDates.length} AO suivi{missingDates.length > 1 ? "s" : ""} sans date de début/fin, donc absent{missingDates.length > 1 ? "s" : ""} du planning : {missingDates.map((a) => a.titre).join(", ")}.
        </p>
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   AUTHENTIFICATION
--------------------------------------------------------- */

// Identifiant d'organisation utilisé en phase 1 (usage interne Co-Concept).
// En phase 3 (revente à d'autres paysagistes), chaque entreprise cliente
// aura son propre orgId au lieu de cette valeur fixe.
const ORG_ID = "co-concept";

function LoginScreen() {
  const [mode, setMode] = useState("login");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    setBusy(true);
    try {
      if (mode === "login") {
        await signInWithEmailAndPassword(auth, email, password);
      } else {
        await createUserWithEmailAndPassword(auth, email, password);
      }
    } catch (err) {
      setError(
        err.code === "auth/invalid-credential"
          ? "E-mail ou mot de passe incorrect."
          : err.code === "auth/email-already-in-use"
          ? "Un compte existe déjà avec cet e-mail — connectez-vous plutôt."
          : err.code === "auth/weak-password"
          ? "Le mot de passe doit faire au moins 6 caractères."
          : "Une erreur est survenue : " + err.message
      );
    } finally {
      setBusy(false);
    }
  };

  return (
    <div
      className="min-h-screen flex items-center justify-center p-6"
      style={{ background: TOKENS.paper }}
    >
      <style>{FONT_IMPORT}</style>
      <div className="w-full max-w-sm p-6 border" style={{ borderColor: TOKENS.ink, background: "white" }}>
        <img src="/logo-co-concept.png" alt="Co-Concept" style={{ height: 48, width: "auto" }} className="mb-4" />
        <div
          className="text-[10px] uppercase tracking-widest mb-1"
          style={{ color: TOKENS.moss, fontFamily: "'JetBrains Mono', monospace", letterSpacing: "0.15em" }}
        >
          Outil de réponse aux appels d'offres
        </div>
        <h1 className="text-2xl mb-6" style={{ color: TOKENS.ink, fontFamily: "'Fraunces', serif", fontWeight: 600 }}>
          AO Paysage
        </h1>

        <form onSubmit={submit} className="flex flex-col gap-3">
          <div>
            <label className="text-xs mb-1 block" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
              E-mail
            </label>
            <input
              type="email"
              required
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              className="w-full p-2 text-sm border outline-none"
              style={{ borderColor: TOKENS.line, fontFamily: "'Inter', sans-serif", color: TOKENS.ink }}
            />
          </div>
          <div>
            <label className="text-xs mb-1 block" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
              Mot de passe
            </label>
            <input
              type="password"
              required
              minLength={6}
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full p-2 text-sm border outline-none"
              style={{ borderColor: TOKENS.line, fontFamily: "'Inter', sans-serif", color: TOKENS.ink }}
            />
          </div>

          {error && (
            <div className="text-xs p-2" style={{ background: TOKENS.rustDim, color: TOKENS.rust, fontFamily: "'Inter', sans-serif" }}>
              {error}
            </div>
          )}

          <button
            type="submit"
            disabled={busy}
            className="flex items-center justify-center gap-2 text-sm py-2 mt-1"
            style={{ background: TOKENS.ink, color: TOKENS.paper, fontFamily: "'Inter', sans-serif", fontWeight: 600 }}
          >
            {busy && <Loader2 size={14} className="animate-spin" />}
            {mode === "login" ? "Se connecter" : "Créer mon compte"}
          </button>
        </form>

        <button
          onClick={() => {
            setError("");
            setMode((m) => (m === "login" ? "signup" : "login"));
          }}
          className="text-xs mt-4 underline"
          style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}
        >
          {mode === "login" ? "Pas encore de compte ? En créer un" : "Déjà un compte ? Se connecter"}
        </button>
      </div>
    </div>
  );
}

function AppContent({ user }) {
  const [activeTab, setActiveTab] = useState("veille");
  const [followed, setFollowed] = useState([]);
  const [statuts, setStatuts] = useState({});
  const [followedDates, setFollowedDates] = useState({});
  const [platforms, setPlatforms] = useState(DEFAULT_PLATFORMS);
  const [openProject, setOpenProject] = useState(null);
  const [projectData, setProjectData] = useState({});
  const [mainLoaded, setMainLoaded] = useState(false);
  const [mainError, setMainError] = useState("");
  const [aos, setAOs] = useState([]);
  const [aosLoaded, setAOsLoaded] = useState(false);
  const seededRef = useRef(false);
  const projectSubs = useRef({});
  const [googleToken, setGoogleToken] = useState(null);
  const [googleExpiry, setGoogleExpiry] = useState(0);
  const [googleError, setGoogleError] = useState("");
  const googleConnected = !!googleToken && Date.now() < googleExpiry;

  const connectGoogle = () => {
    setGoogleError("");
    if (!window.google?.accounts?.oauth2) {
      setGoogleError("Le service Google n'est pas encore chargé, réessayez dans un instant.");
      return;
    }
    const clientId = import.meta.env.VITE_GOOGLE_CLIENT_ID;
    if (!clientId) {
      setGoogleError("Aucun identifiant Google Client configuré (VITE_GOOGLE_CLIENT_ID).");
      return;
    }
    const tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: "https://www.googleapis.com/auth/calendar.events",
      callback: (resp) => {
        if (resp.error) {
          setGoogleError("Connexion refusée : " + resp.error);
          return;
        }
        setGoogleToken(resp.access_token);
        setGoogleExpiry(Date.now() + resp.expires_in * 1000);
      },
    });
    tokenClient.requestAccessToken();
  };

  const TODAY = "2026-08-06";
  const mainRef = doc(db, "organisations", ORG_ID, "app", "main");
  const aosColRef = collection(db, "organisations", ORG_ID, "aos");

  // Tous les AO (démo + ajoutés manuellement) vivent dans Firestore, pour
  // pouvoir être modifiés/supprimés/archivés réellement. Au tout premier
  // lancement (collection vide), on y recopie les AO de démonstration en
  // conservant leurs identifiants (AO-2591, etc.) pour rester compatibles
  // avec les fiches projet déjà préparées.
  useEffect(() => {
    const unsub = onSnapshot(
      aosColRef,
      (snap) => {
        setAOs(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
        setAOsLoaded(true);
        if (snap.empty && !seededRef.current) {
          seededRef.current = true;
          MOCK_AO.forEach(({ id, ...rest }) => {
            setDoc(doc(aosColRef, id), { ...rest, historique: false }).catch(console.error);
          });
        }
      },
      (err) => console.error("Erreur de synchronisation Firestore (aos) :", err)
    );
    return unsub;
  }, []);

  const onAddAO = (data) => addDoc(aosColRef, { ...data, historique: false }).catch(console.error);
  const onUpdateAO = (id, patch) => setDoc(doc(aosColRef, id), patch, { merge: true }).catch(console.error);
  const onDeleteAO = (id) => deleteDoc(doc(aosColRef, id)).catch(console.error);
  const onArchiveAO = (id) => onUpdateAO(id, { historique: true });
  const onRestoreAO = (id) => onUpdateAO(id, { historique: false });

  const activeAOs = useMemo(() => aos.filter((a) => !a.historique), [aos]);
  const historiqueAOs = useMemo(() => aos.filter((a) => a.historique), [aos]);

  // Charge le document principal (AO suivis, statuts, plateformes) et
  // reste synchronisé en temps réel — utile si plusieurs personnes
  // utilisent l'app en même temps.
  useEffect(() => {
    const unsub = onSnapshot(
      mainRef,
      (snap) => {
        if (snap.exists()) {
          const data = snap.data();
          setFollowed(data.followed ?? []);
          setStatuts(data.statuts ?? {});
          setFollowedDates(data.followedDates ?? {});
          setPlatforms(data.platforms ?? DEFAULT_PLATFORMS);
        } else {
          setDoc(mainRef, {
            followed: ["AO-2591"],
            statuts: { "AO-2591": "En préparation" },
            followedDates: { "AO-2591": TODAY },
            platforms: DEFAULT_PLATFORMS,
          });
        }
        setMainLoaded(true);
      },
      (err) => setMainError(err.message)
    );
    return unsub;
  }, []);

  const pushMain = (patch) => setDoc(mainRef, patch, { merge: true }).catch(console.error);

  const onFollow = (id) => {
    setFollowed((f) => {
      const next = f.includes(id) ? f : [...f, id];
      pushMain({ followed: next });
      return next;
    });
    setStatuts((s) => {
      const next = s[id] ? s : { ...s, [id]: "Repéré" };
      pushMain({ statuts: next });
      return next;
    });
    setFollowedDates((d) => {
      const next = d[id] ? d : { ...d, [id]: TODAY };
      pushMain({ followedDates: next });
      return next;
    });
  };
  const onUnfollow = (id) =>
    setFollowed((f) => {
      const next = f.filter((x) => x !== id);
      pushMain({ followed: next });
      return next;
    });
  const onChangeStatut = (id, val) =>
    setStatuts((s) => {
      const next = { ...s, [id]: val };
      pushMain({ statuts: next });
      return next;
    });

  // Abonne (une seule fois par projet) le document Firestore correspondant.
  // Si le projet n'existe pas encore côté serveur, on le crée avec les
  // valeurs par défaut générées localement.
  const ensureProject = (ao) => {
    if (projectSubs.current[ao.id]) return;
    const ref = doc(db, "organisations", ORG_ID, "projects", ao.id);
    const unsub = onSnapshot(
      ref,
      (snap) => {
        if (snap.exists()) {
          const normalized = withProjectDefaults(ao, snap.data());
          setProjectData((pd) => ({ ...pd, [ao.id]: normalized }));
          // Si des champs manquaient (fiche créée avant l'ajout des pièces
          // ou de la mémoire technique), on les enregistre pour de bon.
          setDoc(ref, normalized, { merge: true }).catch(console.error);
        } else {
          const initial = getProjectDetail(ao, followedDates[ao.id] ?? TODAY);
          setDoc(ref, initial).catch(console.error);
          setProjectData((pd) => ({ ...pd, [ao.id]: initial }));
        }
      },
      (err) => console.error("Erreur de synchronisation Firestore (projet " + ao.id + ") :", err)
    );
    projectSubs.current[ao.id] = unsub;
  };
  const onOpenProject = (ao) => {
    ensureProject(ao);
    setOpenProject(ao);
  };
  const updateProject = (aoId, updater) =>
    setProjectData((pd) => {
      if (!pd[aoId]) return pd;
      const next = updater(pd[aoId]);
      setDoc(doc(db, "organisations", ORG_ID, "projects", aoId), next).catch(console.error);
      return { ...pd, [aoId]: next };
    });

  const onUpdateCA = (aoId, monthKey, value) =>
    updateProject(aoId, (d) => ({
      ...d,
      finance: { ...d.finance, caByMonth: { ...d.finance?.caByMonth, [monthKey]: value } },
    }));
  const onUpdateHours = (aoId, memberId, monthKey, value) =>
    updateProject(aoId, (d) => ({
      ...d,
      finance: {
        ...d.finance,
        hours: {
          ...d.finance?.hours,
          [memberId]: { ...d.finance?.hours?.[memberId], [monthKey]: value },
        },
      },
    }));

  const onTogglePlatform = (id) =>
    setPlatforms((ps) => {
      const next = ps.map((p) => (p.id === id ? { ...p, active: !p.active } : p));
      pushMain({ platforms: next });
      return next;
    });
  const onRemovePlatform = (id) =>
    setPlatforms((ps) => {
      const next = ps.filter((p) => p.id !== id);
      pushMain({ platforms: next });
      return next;
    });
  const onAddPlatform = (p) =>
    setPlatforms((ps) => {
      const next = [...ps, p];
      pushMain({ platforms: next });
      return next;
    });

  if (mainError) {
    return (
      <div className="min-h-screen flex items-center justify-center p-6" style={{ background: TOKENS.paper }}>
        <style>{FONT_IMPORT}</style>
        <div className="max-w-sm text-sm p-4 border" style={{ borderColor: TOKENS.rust, background: TOKENS.rustDim, color: TOKENS.rust, fontFamily: "'Inter', sans-serif" }}>
          Impossible de charger les données : {mainError}
          <br />
          Vérifiez que les règles de sécurité Firestore sont bien publiées.
        </div>
      </div>
    );
  }

  if (!mainLoaded || !aosLoaded) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: TOKENS.paper }}>
        <style>{FONT_IMPORT}</style>
        <div className="flex items-center gap-2" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
          <Loader2 size={16} className="animate-spin" /> Chargement des données…
        </div>
      </div>
    );
  }

  return (
    <div style={{ minHeight: "100%", background: TOKENS.paper }}>
      <style>{FONT_IMPORT}</style>
      <Cartouche activeTab={activeTab} />
      <div className="grid grid-cols-12">
        <div className="col-span-12 sm:col-span-2 border-r" style={{ borderColor: TOKENS.line }}>
          <NavItem icon={LayoutDashboard} label="Vue d'ensemble" active={activeTab === "ensemble"} onClick={() => setActiveTab("ensemble")} />
          <NavItem icon={Radar} label="Veille" active={activeTab === "veille"} onClick={() => setActiveTab("veille")} />
          <NavItem icon={KanbanSquare} label="Suivi" active={activeTab === "suivi"} onClick={() => setActiveTab("suivi")} />
          <NavItem icon={FileText} label="Rédaction" active={activeTab === "redaction"} onClick={() => setActiveTab("redaction")} />
          <NavItem icon={Wallet} label="Suivi financier" active={activeTab === "finances"} onClick={() => setActiveTab("finances")} />
          <NavItem icon={Globe} label="Sources" active={activeTab === "sources"} onClick={() => setActiveTab("sources")} />
          <NavItem icon={Archive} label="Historique" active={activeTab === "historique"} onClick={() => setActiveTab("historique")} />
          <div className="px-4 py-4 mt-2 flex items-center gap-2" style={{ color: TOKENS.inkSoft }}>
            <Trophy size={13} />
            <span className="text-[11px]" style={{ fontFamily: "'JetBrains Mono', monospace" }}>
              {followed.length} AO suivi{followed.length > 1 ? "s" : ""}
            </span>
          </div>
          <div className="px-4 py-2 flex flex-col gap-1" style={{ borderTop: `1px solid ${TOKENS.line}` }}>
            <span className="text-[10px] truncate" style={{ color: TOKENS.inkSoft, fontFamily: "'JetBrains Mono', monospace" }}>
              {user.email}
            </span>
            <button
              onClick={() => signOut(auth)}
              className="flex items-center gap-1.5 text-[11px]"
              style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}
            >
              <LogOut size={12} /> Se déconnecter
            </button>
          </div>
          <div className="px-4 py-2" style={{ borderTop: `1px solid ${TOKENS.line}` }}>
            {googleConnected ? (
              <span className="flex items-center gap-1.5 text-[11px]" style={{ color: TOKENS.moss, fontFamily: "'Inter', sans-serif" }}>
                <Check size={12} /> Google Agenda connecté
              </span>
            ) : (
              <button
                onClick={connectGoogle}
                className="flex items-center gap-1.5 text-[11px]"
                style={{ color: TOKENS.blue, fontFamily: "'Inter', sans-serif" }}
              >
                <CalendarDays size={12} /> Connecter mon Google Agenda
              </button>
            )}
            {googleError && (
              <div className="text-[10px] mt-1" style={{ color: TOKENS.rust, fontFamily: "'Inter', sans-serif" }}>{googleError}</div>
            )}
          </div>
        </div>
        <div className="col-span-12 sm:col-span-10">
          {activeTab === "ensemble" && (
            <VueEnsemble aos={activeAOs} followed={followed} projectData={projectData} ensureProject={ensureProject} />
          )}
          {activeTab === "veille" && (
            <Veille
              aos={activeAOs}
              followed={followed}
              onFollow={onFollow}
              onAddAO={onAddAO}
              onUpdateAO={onUpdateAO}
              onDeleteAO={onDeleteAO}
              onArchiveAO={onArchiveAO}
              platforms={platforms}
              googleConnected={googleConnected}
              googleToken={googleToken}
              connectGoogle={connectGoogle}
            />
          )}
          {activeTab === "redaction" && (
            <Redaction aos={activeAOs} followed={followed} projectData={projectData} ensureProject={ensureProject} updateProject={updateProject} />
          )}
          {activeTab === "suivi" && (
            <Suivi
              aos={activeAOs}
              followed={followed}
              statuts={statuts}
              onChangeStatut={onChangeStatut}
              onUnfollow={onUnfollow}
              onOpenProject={onOpenProject}
            />
          )}
          {activeTab === "finances" && (
            <Finances
              followedAOs={activeAOs.filter((a) => followed.includes(a.id))}
              projectData={projectData}
              ensureProject={ensureProject}
              onUpdateCA={onUpdateCA}
              onUpdateHours={onUpdateHours}
            />
          )}
          {activeTab === "sources" && (
            <Sources
              platforms={platforms}
              onToggle={onTogglePlatform}
              onRemove={onRemovePlatform}
              onAdd={onAddPlatform}
            />
          )}
          {activeTab === "historique" && <Historique aos={historiqueAOs} onRestoreAO={onRestoreAO} />}
        </div>

      </div>
      {openProject && projectData[openProject.id] && (
        <ProjectDetail
          key={openProject.id}
          ao={openProject}
          detail={projectData[openProject.id]}
          onClose={() => setOpenProject(null)}
          googleConnected={googleConnected}
          googleToken={googleToken}
          connectGoogle={connectGoogle}
          onUpdateTask={(taskId, patch) =>
            updateProject(openProject.id, (d) => ({
              ...d,
              tasks: d.tasks.map((t) => (t.id === taskId ? { ...t, ...patch } : t)),
            }))
          }
          onAddTask={() =>
            updateProject(openProject.id, (d) => ({
              ...d,
              tasks: [
                ...d.tasks,
                {
                  id: "t-" + Date.now(),
                  label: "Nouvelle tâche",
                  assigneeIds: d.team[0] ? [d.team[0].id] : [],
                  start: d.rangeStart,
                  end: d.rangeEnd,
                },
              ],
            }))
          }
          onRemoveTask={(taskId) =>
            updateProject(openProject.id, (d) => ({ ...d, tasks: d.tasks.filter((t) => t.id !== taskId) }))
          }
          onUpdateMember={(memberId, patch) =>
            updateProject(openProject.id, (d) => ({
              ...d,
              team: d.team.map((m) => (m.id === memberId ? { ...m, ...patch } : m)),
            }))
          }
          onAddMember={() =>
            updateProject(openProject.id, (d) => ({
              ...d,
              team: [...d.team, { id: "mb-" + Date.now(), name: "", poste: "", mandataire: false }],
            }))
          }
          onRemoveMember={(memberId) =>
            updateProject(openProject.id, (d) => ({ ...d, team: d.team.filter((m) => m.id !== memberId) }))
          }
          onSetMandataire={(memberId) =>
            updateProject(openProject.id, (d) => ({
              ...d,
              team: d.team.map((m) => ({ ...m, mandataire: m.id === memberId })),
            }))
          }
          onUpdateLinks={(patch) =>
            updateProject(openProject.id, (d) => ({ ...d, links: { ...d.links, ...patch } }))
          }
          onAddMeeting={() => {
            const id = "r-" + Date.now();
            updateProject(openProject.id, (d) => ({
              ...d,
              meetings: [
                ...d.meetings,
                { id, title: "Nouvelle réunion", date: d.rangeStart, time: "09:00", attendees: [] },
              ],
            }));
            return id;
          }}
          onUpdateMeeting={(meetingId, patch) =>
            updateProject(openProject.id, (d) => ({
              ...d,
              meetings: d.meetings.map((r) => (r.id === meetingId ? { ...r, ...patch } : r)),
            }))
          }
          onRemoveMeeting={(meetingId) =>
            updateProject(openProject.id, (d) => ({ ...d, meetings: d.meetings.filter((r) => r.id !== meetingId) }))
          }
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------
   APP (authentification)
--------------------------------------------------------- */

export default function App() {
  const [user, setUser] = useState(undefined); // undefined = en cours de vérification, null = déconnecté
  const [orgReady, setOrgReady] = useState(false);
  const [orgError, setOrgError] = useState("");

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (u) => setUser(u));
    return unsub;
  }, []);

  // Chaque personne qui se connecte pour la première fois obtient un
  // document /users/{uid} rattaché à l'organisation Co-Concept — c'est ce
  // qui lui donne le droit de lire/écrire les données de l'organisation
  // selon les règles de sécurité Firestore. On attend que ce document soit
  // confirmé avant de charger le reste de l'app, pour éviter toute lecture
  // refusée par les règles de sécurité.
  useEffect(() => {
    if (!user) {
      setOrgReady(false);
      return;
    }
    const uref = doc(db, "users", user.uid);
    getDoc(uref)
      .then((snap) => (snap.exists() ? null : setDoc(uref, { orgId: ORG_ID, email: user.email })))
      .then(() => setOrgReady(true))
      .catch((err) => setOrgError(err.message));
  }, [user]);

  if (user === undefined) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: TOKENS.paper }}>
        <style>{FONT_IMPORT}</style>
        <Loader2 size={20} className="animate-spin" style={{ color: TOKENS.inkSoft }} />
      </div>
    );
  }

  if (!user) {
    return <LoginScreen />;
  }

  if (orgError) {
    return (
      <div className="min-h-screen flex items-center justify-center p-6" style={{ background: TOKENS.paper }}>
        <style>{FONT_IMPORT}</style>
        <div className="max-w-sm text-sm p-4 border" style={{ borderColor: TOKENS.rust, background: TOKENS.rustDim, color: TOKENS.rust, fontFamily: "'Inter', sans-serif" }}>
          Impossible d'accéder à Firestore : {orgError}
          <br />
          Vérifiez que les règles de sécurité sont bien publiées.
        </div>
      </div>
    );
  }

  if (!orgReady) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: TOKENS.paper }}>
        <style>{FONT_IMPORT}</style>
        <div className="flex items-center gap-2" style={{ color: TOKENS.inkSoft, fontFamily: "'Inter', sans-serif" }}>
          <Loader2 size={16} className="animate-spin" /> Préparation de votre compte…
        </div>
      </div>
    );
  }

  return <AppContent user={user} />;
}
