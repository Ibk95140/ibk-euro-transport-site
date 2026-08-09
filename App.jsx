import React, { useState, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import 'jspdf-autotable';
import { supabase } from './lib/supabase';
import Login from './components/Login';
import { LOGO_IBK_BASE64 } from './assets/logoBase64';
import { SOCIETE } from './assets/societe';

// ===== CONSTANTES =====
const STATUS_MAP = {
  planifiee: { cls: 'bg-sky-500/15 text-sky-700', label: 'Planifié' },
  charge: { cls: 'bg-amber-400/15 text-amber-700', label: 'Chargé 📦' },
  en_cours: { cls: 'bg-[#d4006e]/15 text-[#d4006e]', label: 'En cours 🚛' },
  livree: { cls: 'bg-emerald-500/15 text-emerald-700', label: 'Délivré ✓' },
  annule: { cls: 'bg-gray-400/15 text-gray-600', label: 'Annulé ✕' },
  litige: { cls: 'bg-red-500/15 text-red-700', label: 'Litige ⚠' },
};

// Statuts documentaires : 100% calculés automatiquement à partir des dates
// de validité saisies par le sous-traitant lors du dépôt — aucune
// validation humaine, aucun choix manuel. 4 états possibles :
// Manquant (pas de document ou pas de date de fin), Valide, Expire bientôt
// (à moins de 30 jours de l'échéance) et Expiré (date de fin dépassée).
const DOC_BADGE = {
  ok: { cls: 'bg-emerald-500/15 text-emerald-700', label: 'Valide' },
  expire_bientot: { cls: 'bg-amber-400/15 text-amber-700', label: '⚠ Expire bientôt' },
  expire: { cls: 'bg-red-500/15 text-red-700', label: 'Expiré' },
  manquant: { cls: 'bg-gray-200 text-gray-600', label: 'Manquant' },
};

const SEUIL_EXPIRE_BIENTOT_JOURS = 30;

// Calcule le statut d'un document uniquement à partir de sa présence et de
// sa date de fin de validité — jamais d'intervention humaine, jamais de
// valeur stockée manuellement.
function computeEffectiveStatut(dateFin, hasUrl) {
  if (!hasUrl || !dateFin) return 'manquant';
  const todayStr = new Date().toISOString().slice(0, 10);
  if (todayStr > dateFin) return 'expire';
  const today = new Date();
  const fin = new Date(dateFin + 'T00:00:00');
  const diffJours = Math.ceil((fin - today) / (1000 * 60 * 60 * 24));
  if (diffJours <= SEUIL_EXPIRE_BIENTOT_JOURS) return 'expire_bientot';
  return 'ok';
}

const FACTURE_STATUS_MAP = {
  emise: { cls: 'bg-[#0d1b2a]/15 text-[#0d1b2a]', label: 'Émise' },
  payee: { cls: 'bg-emerald-500/15 text-emerald-700', label: 'Payée ✓' },
  retard: { cls: 'bg-red-500/15 text-red-700', label: 'En retard' },
};

// ===== MENTION LÉGALE TVA (transport intracommunautaire) =====
// Ce TMS ne gère que l'affrètement intracommunautaire : aucune TVA n'est
// appliquée sur les factures (client et transporteur), le montant facturé
// correspond exactement à la somme convenue. Le mécanisme applicable est
// en principe l'autoliquidation (reverse charge) côté preneur.
// ⚠️ La formulation exacte ci-dessous est une mention par défaut : à faire
// valider avec l'expert-comptable d'IBK EURO AFRIQUE avant usage définitif.
const MENTION_TVA = "TVA non applicable - Autoliquidation par le preneur - Article 283-2 du CGI / Article 44 de la Directive 2006/112/CE";

function factureFromRow(row) {
  return {
    id: row.id,
    numero: row.numero,
    sens: row.sens,
    missionRef: row.mission_ref,
    tiers: row.tiers,
    date: row.date_facture,
    echeance: row.date_echeance,
    montantHT: Number(row.montant_ht) || 0,
    statut: row.statut,
    notes: row.notes,
    datePaiement: row.date_paiement,
    notifComptable: row.notif_comptable,
    factureUrl: row.facture_url,
    cmrUrl: row.cmr_url,
    valideComptable: row.valide_comptable,
    nbJoursPaiement: row.nb_jours_paiement || 30,
  };
}


function margeClass(pct) {
  if (pct < 8) return 'bg-red-500/15 text-red-700';
  if (pct < 15) return 'bg-amber-400/15 text-amber-700';
  return 'bg-emerald-500/15 text-emerald-700';
}

// CORRECTION : on remplace les espaces spéciales (insécable / fine insécable)
// générées par toLocaleString par une espace normale, car la police utilisée
// dans les PDF (jsPDF/Helvetica) affiche ces caractères comme un "/".
function fmtEUR(n) {
  return (n || 0).toLocaleString('fr-FR').replace(/\u202F|\u00A0/g, ' ') + ' €';
}
function fmtDate(iso) {
  if (!iso) return '—';
  const d = new Date(iso + 'T00:00:00');
  return d.toLocaleDateString('fr-FR');
}

// ===== ADAPTATEURS SUPABASE =====
// Ces fonctions traduisent les noms de colonnes de la base de données
// (snake_case, ex: date_mission) vers le format utilisé par l'interface
// (camelCase, ex: date), et inversement.

function missionFromRow(row) {
  return {
    ref: row.ref,
    flux: row.flux,
    source: row.source,
    partenaire: row.partenaire,
    depart: row.depart,
    dest: row.destination,
    date: row.date_mission,
    dateLivraison: row.date_livraison,
    vendu: Number(row.vendu) || 0,
    paye: Number(row.paye) || 0,
    statut: row.statut,
    affreteur: row.affreteur,
    affreteurEmail: row.affreteur_email,
    affreteurTel: row.affreteur_tel,
    transporteurAdresse: row.transporteur_adresse,
    transporteurContact: row.transporteur_contact,
    transporteurContactEmail: row.transporteur_contact_email,
    transporteurContactTel: row.transporteur_contact_tel,
    plaque: row.plaque,
    contactChargement: row.contact_chargement,
    heureChargement: row.heure_chargement,
    heureChargementFin: row.heure_chargement_fin,
    contactLivraison: row.contact_livraison,
    heureLivraison: row.heure_livraison,
    heureLivraisonFin: row.heure_livraison_fin,
    cpCharge: row.cp_charge,
    paysCharge: row.pays_charge,
    cpLivre: row.cp_livre,
    paysLivre: row.pays_livre,
    marchandise: row.marchandise,
    poids: row.poids,
    volume: row.volume,
    ldm: row.ldm,
    exchangePalettes: !!row.exchange_palettes,
    delaiPaiement: row.delai_paiement,
    refClient: row.ref_client,
    ordreClientUrl: row.ordre_client_url,
    cmrUrl: row.cmr_url,
  };
}

function stFromRow(row) {
  return {
    id: row.id,
    nom: row.nom,
    siret: row.siret,
    contact: row.contact,
    tel: row.telephone,
    email: row.email,
    numeroTeleroute: row.numero_teleroute,
    numeroTimocom: row.numero_timocom,
    flotte: row.flotte,
    zone: row.zone,
    note: row.note,
    typeEntreprise: row.type_entreprise || 'transporteur',
    siretVerifie: row.siret_verifie,
    siretStatut: row.siret_statut,
    siretNomOfficiel: row.siret_nom_officiel,
    siretDerniereVerif: row.siret_derniere_verif,
    numeroTva: row.numero_tva,
    tvaVerifie: row.tva_verifie,
    tvaStatut: row.tva_statut,
    tvaPays: row.tva_pays,
    tvaNomOfficiel: row.tva_nom_officiel,
  };
}

function attestationFromRow(row) {
  return {
    id: row.id,
    stId: row.sous_traitant_id,
    licence: row.licence,
    licenceDateDebut: row.licence_date_debut,
    licenceDate: row.licence_date,
    assurance: row.assurance,
    assuranceDateDebut: row.assurance_date_debut,
    assuranceDate: row.assurance_date,
    kbis: row.kbis,
    kbisDateDebut: row.kbis_date_debut,
    kbisDate: row.kbis_date,
    chauffeurNom: row.chauffeur_nom,
    chauffeurTel: row.chauffeur_tel,
    publicToken: row.public_token,
    licenceUrl: row.licence_url,
    assuranceUrl: row.assurance_url,
    kbisUrl: row.kbis_url,
  };
}

function donneurFromRow(row) {
  return {
    id: row.id,
    nom: row.nom,
    type: row.type,
    delai: row.delai_paiement,
    statut: row.statut,
    numeroClient: row.numero_client,
    numeroTva: row.numero_tva,
    tvaVerifie: row.tva_verifie,
    tvaStatut: row.tva_statut,
    tvaPays: row.tva_pays,
    scoreSolvabilite: row.score_solvabilite,
    scoreDetails: row.score_details,
    scoreDate: row.score_date,
  };
}

// Demandes reçues via le site vitrine IBK Euro Transport (formulaire de
// contact ou "IBK Connect" : transporteur / commissionnaire / chargeur).
function demandeDevisFromRow(row) {
  return {
    id: row.id,
    type: row.type_demande,
    nom: row.nom,
    entreprise: row.entreprise,
    email: row.email,
    telephone: row.telephone,
    siret: row.siret,
    profil: row.profil,
    message: row.message,
    details: row.details || {},
    statut: row.statut || 'nouveau',
    date: row.created_at,
  };
}

// ===== TOAST =====
function Toast({ message }) {
  if (!message) return null;
  return (
    <div className="fixed bottom-6 right-6 z-50 bg-[#ffffff] border border-[#d4006e] text-[#1c2733] px-5 py-3 rounded-lg text-sm font-semibold shadow-xl animate-[fadeIn_0.2s_ease]">
      {message}
    </div>
  );
}

function AppShell({ userEmail, onLogout, role, poste, profil }) {
  const isAdmin = role === 'admin';
  // Droits d'accès personnalisés : si un administrateur a explicitement
  // défini une liste de modules pour cet utilisateur (indépendamment de son
  // poste), elle prévaut sur les permissions par défaut du poste. Sinon,
  // on retombe sur les permissions standards associées au poste.
  const customPerms = Array.isArray(profil?.permissions_custom) ? profil.permissions_custom : null;
  function canSee(moduleKey, defaultForPoste) {
    if (isAdmin) return true;
    if (customPerms) return customPerms.includes(moduleKey);
    return defaultForPoste;
  }
  // Permissions par poste (ou personnalisées si définies)
  const canSeeMarges = canSee('marges', poste === 'superviseur' || poste === 'commercial');
  const canSeeFacturation = canSee('facturation', poste === 'comptable');
  const canSeeAffectation = canSee('affectation', poste === 'commercial' || poste === 'exploitant');
  const canSeeAxes = canSee('historique-axes', poste === 'superviseur');
  const canCreateOrdre = isAdmin || poste === 'exploitant' || poste === 'commercial';
  const canManageSousTraitants = canSee('sous-traitants', poste === 'exploitant');
  const canSeeClients = canSee('donneurs', poste === 'commercial' || poste === 'exploitant');
  const canSeeDemandes = canSee('demandes-devis', poste === 'commercial' || poste === 'exploitant');
  const [loading, setLoading] = useState(true);
  const [missions, setMissions] = useState([]);
  const [sousTraitants, setSousTraitants] = useState([]);
  const [attestations, setAttestations] = useState([]);
  const [donneurs, setDonneurs] = useState([]);
  const [demandesDevis, setDemandesDevis] = useState([]);
  const [factures, setFactures] = useState([]);
  const [view, setView] = useState('dashboard');
  const [toast, setToast] = useState('');
  const [flowFilter, setFlowFilter] = useState('all');
  const [statusFilter, setStatusFilter] = useState('all');
  const [factureFilter, setFactureFilter] = useState('all');

  const [stModal, setStModal] = useState(false);
  const [donneurModal, setDonneurModal] = useState(false);
  const [factureModal, setFactureModal] = useState(false);
  const [attestationModal, setAttestationModal] = useState(false);
  const [docModal, setDocModal] = useState(null); // mission courante pour la modale "Documents" (ordre client + CMR)
  const [docUploading, setDocUploading] = useState(false);

  const [stForm, setStForm] = useState({ id: null, nom: '', siret: '', contact: '', tel: '', email: '', numeroTeleroute: '', numeroTimocom: '', flotte: '', zone: '', typeEntreprise: 'transporteur', numeroTva: '' });
  const [stSiretCheck, setStSiretCheck] = useState(null);
  const [tvaCheck, setTvaCheck] = useState(null); // vérification TVA en direct dans les modals // { loading, statut, nom } — aperçu dans la modale, avant sauvegarde
  // Méthode de vérification sélectionnée dans le bouton unique "Vérifier" :
  // TVA (méthode principale, valable pour toute entreprise européenne) ou
  // SIRET (entreprises françaises). Remplace l'ancien bouton "Vérifier SIRET".
  const [stVerifMethod, setStVerifMethod] = useState('tva');
  const [donneurVerifMethod, setDonneurVerifMethod] = useState('tva');
  const [donneurForm, setDonneurForm] = useState({ nom: '', type: 'Commissionnaire de transport', delai: '', siret: '', numeroTva: '' });
  const [factureForm, setFactureForm] = useState({
    sens: 'client', missionRef: '', tiers: '', date: '', echeance: '', montantHT: '', notes: '',
  });
  const [attestationForm, setAttestationForm] = useState({
    id: null, licenceDateDebut: '', licenceDate: '',
    assuranceDateDebut: '', assuranceDate: '',
    kbisDateDebut: '', kbisDate: '', chauffeurNom: '', chauffeurTel: '',
  });


  // cotation
  const [cotPaysDepart, setCotPaysDepart] = useState('France');
  const [cotCpDepart, setCotCpDepart] = useState('');
  const [cotPaysArrivee, setCotPaysArrivee] = useState('France');
  const [cotCpArrivee, setCotCpArrivee] = useState('');
  const [cotKm, setCotKm] = useState('');
  const [cotTypeCamion, setCotTypeCamion] = useState('Tautliner');
  const [cotPeage, setCotPeage] = useState('0');

  // ===== ORDRE DE TRANSPORT (unique formulaire, remplace l'ancien "+ Nouvel
  // ordre") =====
  // Ce formulaire est désormais le SEUL endroit où une mission (ordre
  // transporteur) est créée dans le TMS : le remplir et cliquer sur
  // "Créer la mission et télécharger le PDF" crée automatiquement la
  // mission (visible ensuite dans Ordres/Marges/Facturation) ET génère le
  // document PDF. Si `otExistingRef` est renseigné, on est en mode
  // "réimpression" d'une mission déjà créée (pas de nouvelle insertion,
  // juste régénération du PDF avec les données actuelles).
  const [otExistingRef, setOtExistingRef] = useState(null);
  const [lv, setLv] = useState({
    ref: '', date: '', dateLivraison: '',
    affreteur: '', affreteurEmail: '', affreteurTel: '',
    transporteur: '', transporteurAdresse: '', transporteurContact: '', transporteurContactEmail: '', transporteurContactTel: '',
    plaque: '',
    lieuCharge: '', cpCharge: '', paysCharge: 'France',
    lieuLivre: '', cpLivre: '', paysLivre: 'France',
    clientChargement: '', heureChargementDebut: '', heureChargementFin: '',
    clientLivraison: '', heureLivraisonDebut: '', heureLivraisonFin: '',
    marchandise: '', poids: '', ldm: '', exchangePalettes: false,
    // prixClient : usage strictement interne (Marges/Facturation/Tableau de
    // bord) — ce champ n'est JAMAIS imprimé sur le PDF Ordre de transport
    // remis au transporteur (voir exportOTPdf, qui ne le lit jamais).
    prixClient: '', prix: '', delaiPaiement: '30 days end of month',
  });

  function resetLv() {
    setLv({
      ref: '', date: '', dateLivraison: '',
      affreteur: '', affreteurEmail: '', affreteurTel: '',
      transporteur: '', transporteurAdresse: '', transporteurContact: '', transporteurContactEmail: '', transporteurContactTel: '',
      plaque: '',
      lieuCharge: '', cpCharge: '', paysCharge: 'France',
      lieuLivre: '', cpLivre: '', paysLivre: 'France',
      clientChargement: '', heureChargementDebut: '', heureChargementFin: '',
      clientLivraison: '', heureLivraisonDebut: '', heureLivraisonFin: '',
      marchandise: '', poids: '', ldm: '', exchangePalettes: false,
      prixClient: '', prix: '', delaiPaiement: '30 days end of month',
    });
    setOtExistingRef(null);
  }

  const showToast = useCallback((msg) => {
    setToast(msg);
    setTimeout(() => setToast(''), 3000);
  }, []);

  // ===== CHARGEMENT INITIAL DEPUIS SUPABASE =====
  const loadAll = useCallback(async () => {
    const [missionsRes, stRes, atRes, dnRes, fcRes, ddRes] = await Promise.all([
      supabase.from('missions').select('*').order('created_at', { ascending: false }),
      supabase.from('sous_traitants').select('*').order('nom'),
      supabase.from('attestations').select('*'),
      supabase.from('donneurs_ordre').select('*').order('nom'),
      supabase.from('factures').select('*').order('created_at', { ascending: false }),
      supabase.from('demandes_devis').select('*').order('created_at', { ascending: false }),
    ]);

    if (missionsRes.error || stRes.error || atRes.error || dnRes.error || fcRes.error || ddRes.error) {
      console.error(missionsRes.error || stRes.error || atRes.error || dnRes.error || fcRes.error || ddRes.error);
      showToast('⚠️ Erreur de chargement des données');
    }

    setMissions((missionsRes.data || []).map(missionFromRow));
    setSousTraitants((stRes.data || []).map(stFromRow));
    setAttestations((atRes.data || []).map(attestationFromRow));
    setDonneurs((dnRes.data || []).map(donneurFromRow));
    setFactures((fcRes.data || []).map(factureFromRow));
    setDemandesDevis((ddRes.data || []).map(demandeDevisFromRow));
    setLoading(false);
  }, [showToast]);

  useEffect(() => {
    loadAll();

    // Mises à jour en temps réel : si DIA Houdou ajoute une mission,
    // ton écran se met à jour automatiquement sans avoir à recharger.
    const channel = supabase
      .channel('ibk-tms-realtime')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'missions' }, loadAll)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'sous_traitants' }, loadAll)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'attestations' }, loadAll)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'donneurs_ordre' }, loadAll)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'factures' }, loadAll)
      .subscribe();

    return () => {
      supabase.removeChannel(channel);
    };
  }, [loadAll]);

  // ===== QUI EST CONNECTÉ EN CE MOMENT (visible par les admins) =====
  const [onlineUsers, setOnlineUsers] = useState([]);

  useEffect(() => {
    const presenceChannel = supabase.channel('ibk-presence', {
      config: { presence: { key: userEmail } },
    });

    presenceChannel
      .on('presence', { event: 'sync' }, () => {
        const state = presenceChannel.presenceState();
        const users = Object.values(state)
          .flat()
          .map((u) => u.email)
          .filter(Boolean);
        setOnlineUsers([...new Set(users)]);
      })
      .subscribe(async (status) => {
        if (status === 'SUBSCRIBED') {
          await presenceChannel.track({ email: userEmail, online_at: new Date().toISOString() });
        }
      });

    return () => {
      supabase.removeChannel(presenceChannel);
    };
  }, [userEmail]);

  // ===== DERIVED =====
  const outMissions = missions.filter((m) => m.flux === 'out');
  const inMissions = missions.filter((m) => m.flux === 'in');
  const totalOut = outMissions.reduce((s, m) => s + (m.vendu || 0), 0);
  const totalIn = inMissions.reduce((s, m) => s + (m.vendu || 0), 0);
  const totalMarge = outMissions.reduce((s, m) => s + ((m.vendu || 0) - (m.paye || 0)), 0);
  const margeAvgPct = outMissions.length ? (totalMarge / totalOut) * 100 : 0;

  const filteredMissions = missions.filter((m) => {
    if (flowFilter !== 'all' && m.flux !== flowFilter) return false;
    if (statusFilter !== 'all' && m.statut !== statusFilter) return false;
    return true;
  });

  function stNameById(id) {
    const s = sousTraitants.find((x) => x.id === id);
    return s ? s.nom : '—';
  }

  // ===== ACTIONS =====

  // Téléverse un document lié à un ordre reçu (ordre client original ou CMR)
  // dans le bucket dédié, et renvoie le chemin de stockage.
  async function uploadOrdreRecuDoc(ref, file, type) {
    if (!file) return null;
    const path = `${ref}/${type}-${Date.now()}-${file.name}`;
    const { error } = await supabase.storage.from('documents-clients').upload(path, file);
    if (error) {
      console.error(error);
      return null;
    }
    return path;
  }

  // ===== DOCUMENTS D'UN ORDRE REÇU (ordre client original + CMR) =====
  // Permet de téléverser ou remplacer ces deux documents après coup, pour
  // les ordres reçus créés avant réception du mail du client, ou pour
  // corriger un document déjà déposé.
  async function uploadDocForMission(ref, file, type) {
    if (!file) return;
    setDocUploading(true);
    const path = await uploadOrdreRecuDoc(ref, file, type);
    if (!path) {
      setDocUploading(false);
      showToast('⚠️ Erreur lors du téléversement');
      return;
    }
    const column = type === 'ordre-client' ? 'ordre_client_url' : 'cmr_url';
    const { error } = await supabase.from('missions').update({ [column]: path }).eq('ref', ref);
    setDocUploading(false);
    if (error) {
      showToast('⚠️ Erreur lors de l\'enregistrement du document');
      console.error(error);
      return;
    }
    await loadAll();
    setDocModal((prev) => prev ? { ...prev, [type === 'ordre-client' ? 'ordreClientUrl' : 'cmrUrl']: path } : prev);
    showToast(type === 'ordre-client' ? "✅ Ordre du client déposé" : '✅ CMR déposée');
  }

  async function viewClientDocument(path) {
    if (!path) {
      showToast('⚠️ Aucun document déposé pour le moment');
      return;
    }
    const { data, error } = await supabase.storage.from('documents-clients').createSignedUrl(path, 3600);
    if (error || !data) {
      showToast('⚠️ Impossible d\'ouvrir le document');
      console.error(error);
      return;
    }
    window.open(data.signedUrl, '_blank');
  }

  async function createSousTraitant() {
    if (!stForm.nom) {
      showToast('⚠️ Le nom est requis');
      return;
    }

    if (stForm.id) {
      const { error } = await supabase
        .from('sous_traitants')
        .update({ nom: stForm.nom, siret: stForm.siret, contact: stForm.contact, telephone: stForm.tel, email: stForm.email || null, numero_teleroute: stForm.numeroTeleroute || null, numero_timocom: stForm.numeroTimocom || null, flotte: stForm.flotte, zone: stForm.zone, type_entreprise: stForm.typeEntreprise, numero_tva: stForm.numeroTva || null })
        .eq('id', stForm.id);

      if (error) {
        showToast('⚠️ Erreur lors de la modification du sous-traitant');
        console.error(error);
        return;
      }

      await loadAll();
      setStModal(false);
      setStForm({ id: null, nom: '', siret: '', contact: '', tel: '', email: '', numeroTeleroute: '', numeroTimocom: '', flotte: '', zone: '', typeEntreprise: 'transporteur', numeroTva: '' });
      showToast(`✅ ${stForm.nom} mis à jour`);
      return;
    }

    const { data, error } = await supabase
      .from('sous_traitants')
      .insert({ nom: stForm.nom, siret: stForm.siret, contact: stForm.contact, telephone: stForm.tel, email: stForm.email || null, numero_teleroute: stForm.numeroTeleroute || null, numero_timocom: stForm.numeroTimocom || null, flotte: stForm.flotte, zone: stForm.zone, type_entreprise: stForm.typeEntreprise, numero_tva: stForm.numeroTva || null })
      .select()
      .single();

    if (error) {
      showToast('⚠️ Erreur lors de l\'ajout du sous-traitant');
      console.error(error);
      return;
    }

    await supabase.from('attestations').insert({ sous_traitant_id: data.id });

    await loadAll();
    setStModal(false);
    setStForm({ id: null, nom: '', siret: '', contact: '', tel: '', email: '', numeroTeleroute: '', numeroTimocom: '', flotte: '', zone: '', typeEntreprise: 'transporteur', numeroTva: '' });
    showToast(`✅ ${stForm.nom} ajouté à la liste des sous-traitants`);
  }

  function openEditSousTraitant(s) {
    setStForm({ id: s.id, nom: s.nom, siret: s.siret || '', contact: s.contact || '', tel: s.tel || '', email: s.email || '', numeroTeleroute: s.numeroTeleroute || '', numeroTimocom: s.numeroTimocom || '', flotte: s.flotte || '', zone: s.zone || '', typeEntreprise: s.typeEntreprise || 'transporteur', numeroTva: s.numeroTva || '' });
    setStSiretCheck(null);
    setTvaCheck(null);
    // Si l'entreprise a déjà un SIRET mais pas de TVA, on part sur SIRET ;
    // sinon TVA reste la méthode par défaut (recommandée, valable partout).
    setStVerifMethod(!s.numeroTva && s.siret ? 'siret' : 'tva');
    setStModal(true);
  }

  async function deleteSousTraitant(s) {
    const confirmed = window.confirm(`Supprimer définitivement ${s.nom} ? Ses documents et attestations seront aussi supprimés.`);
    if (!confirmed) return;

    const { error } = await supabase.from('sous_traitants').delete().eq('id', s.id);
    if (error) {
      showToast('⚠️ Erreur lors de la suppression');
      console.error(error);
      return;
    }
    await loadAll();
    showToast(`${s.nom} supprimé`);
  }

  async function deleteDonneur(d) {
    const confirmed = window.confirm(`Supprimer définitivement ${d.nom} de la liste des clients ?`);
    if (!confirmed) return;

    const { error } = await supabase.from('donneurs_ordre').delete().eq('id', d.id);
    if (error) {
      showToast('⚠️ Erreur lors de la suppression');
      console.error(error);
      return;
    }
    await loadAll();
    showToast(`${d.nom} supprimé`);
  }

  // ===== DEMANDES DE DEVIS (issues du site vitrine) =====
  async function updateDemandeStatut(id, newStatut) {
    const { error } = await supabase.from('demandes_devis').update({ statut: newStatut }).eq('id', id);
    if (error) {
      showToast('⚠️ Erreur lors de la mise à jour');
      console.error(error);
      return;
    }
    await loadAll();
  }

  async function deleteDemande(d) {
    const confirmed = window.confirm(`Supprimer définitivement cette demande (${d.nom || d.entreprise || 'sans nom'}) ?`);
    if (!confirmed) return;

    const { error } = await supabase.from('demandes_devis').delete().eq('id', d.id);
    if (error) {
      showToast('⚠️ Erreur lors de la suppression');
      console.error(error);
      return;
    }
    await loadAll();
    showToast('Demande supprimée');
  }

  // Pré-remplit le formulaire "Ajouter un client" à partir d'une demande
  // reçue via le site, pour éviter une re-saisie manuelle.
  function convertDemandeEnClient(d) {
    setDonneurForm({
      nom: d.entreprise || d.nom || '',
      type: 'Commissionnaire de transport',
      delai: '',
      siret: d.siret || '',
      numeroTva: '',
    });
    setView('donneurs');
    setTimeout(() => setDonneurModal(true), 50);
    showToast('Formulaire pré-rempli à partir de la demande — vérifie et enregistre');
  }

  // ===== APERÇU SOLVABILITÉ DANS LE FORMULAIRE (sans sauvegarder en base) =====
  async function previewSolvabilite(siret) {
    const siretClean = (siret || '').replace(/\s/g, '');
    if (!siretClean) { showToast('⚠️ Renseigne le SIRET d\'abord'); return; }
    showToast('Calcul en cours…');
    try {
      const res = await fetch(`https://recherche-entreprises.api.gouv.fr/search?q=${siretClean}&per_page=1`);
      const json = await res.json();
      const result = json.results && json.results[0];
      if (!result) { showToast('⚠️ Entreprise introuvable'); return; }
      const siege = result.siege || {};
      const etatAdmin = siege.etat_administratif || result.etat_administratif;
      const dateCreation = result.date_creation;
      const capital = result.capital || 0;
      let score = 15;
      if (etatAdmin === 'A') score += 25;
      if (dateCreation) {
        const annees = Math.floor((Date.now() - new Date(dateCreation)) / (1000 * 60 * 60 * 24 * 365));
        if (annees >= 10) score += 25;
        else if (annees >= 5) score += 15;
        else if (annees >= 2) score += 8;
        else if (annees >= 1) score += 3;
      }
      if (capital >= 100000) score += 20;
      else if (capital >= 50000) score += 15;
      else if (capital >= 10000) score += 10;
      else if (capital > 0) score += 5;
      const eff = parseInt(result.tranche_effectif_salarie || '0') || 0;
      if (eff >= 50) score += 15;
      else if (eff >= 10) score += 10;
      else if (eff >= 3) score += 5;
      else score += 2;
      score = Math.min(100, score);
      const mention = score >= 70 ? '🟢 Solide' : score >= 45 ? '🟡 Correct' : '🔴 Risqué';
      const nom = result.nom_complet || '';
      showToast(`✅ Score : ${score}/100 — ${mention}${nom ? ` (${nom})` : ''}`);
    } catch (e) {
      showToast('⚠️ Erreur lors du calcul');
    }
  }

  // ===== VÉRIFICATION SIRET EN DIRECT DANS LA MODALE (aperçu avant sauvegarde) =====
  async function checkSiretLive() {
    const siretClean = (stForm.siret || '').replace(/\s/g, '');
    if (siretClean.length < 9) {
      showToast('⚠️ SIRET/SIREN trop court pour vérification');
      return;
    }
    setStSiretCheck({ loading: true });
    try {
      const res = await fetch(`https://recherche-entreprises.api.gouv.fr/search?q=${siretClean}&per_page=1`);
      const json = await res.json();
      const result = json.results && json.results[0];

      if (!result) {
        setStSiretCheck({ loading: false, statut: 'introuvable', nom: null });
        return;
      }
      const etatAdmin = result.siege?.etat_administratif || result.etat_administratif;
      const statut = etatAdmin === 'A' ? 'actif' : 'cesse';
      const nomOfficiel = result.nom_complet || result.nom_raison_sociale || null;
      const siege = result.siege || {};
      const adresseParts = [siege.numero_voie, siege.type_voie, siege.libelle_voie].filter(Boolean).join(' ');
      const adresseComplete = [adresseParts, siege.code_postal, siege.libelle_commune].filter(Boolean).join(' ');

      // Le nom du gérant/dirigeant n'est disponible que pour certaines
      // formes juridiques (SARL, SASU...) ; on prend le premier dirigeant
      // listé, le plus souvent le gérant principal.
      let nomDirigeant = null;
      if (Array.isArray(result.dirigeants) && result.dirigeants.length > 0) {
        const d = result.dirigeants[0];
        if (d.nom || d.prenoms) {
          nomDirigeant = [d.qualite, [d.prenoms, d.nom].filter(Boolean).join(' ')].filter(Boolean).join(' — ');
        } else if (d.denomination) {
          nomDirigeant = d.denomination; // cas d'une personne morale dirigeante
        }
      }

      setStSiretCheck({ loading: false, statut, nom: nomOfficiel, adresse: adresseComplete || null, dirigeant: nomDirigeant });

      // Remplit automatiquement la raison sociale et le contact avec les
      // informations officielles. La zone n'est complétée que si elle est
      // encore vide, pour ne pas écraser une description plus large que
      // l'utilisateur aurait déjà saisie (ex : "Île-de-France, National").
      if (nomOfficiel || nomDirigeant || siege.libelle_commune) {
        setStForm((prev) => ({
          ...prev,
          nom: nomOfficiel || prev.nom,
          contact: nomDirigeant || prev.contact,
          zone: prev.zone ? prev.zone : (siege.libelle_commune || prev.zone),
        }));
      }
    } catch (e) {
      console.error(e);
      setStSiretCheck({ loading: false, statut: 'erreur', nom: null });
    }
  }

  // ===== APPEL AU REGISTRE VIES VIA LE RELAIS SUPABASE (Edge Function) =====
  // Le service européen VIES n'autorise pas les appels directs depuis un
  // navigateur (CORS bloqué côté serveur européen) : il faut donc passer par
  // une fonction serveur qui relaie la requête. Voir la fonction Edge
  // "verify-vat" à créer dans Supabase (Edge Functions > Deploy > Via Editor).
  async function callViesRelay(countryCode, vatNumber) {
    const { data, error } = await supabase.functions.invoke('verify-vat', {
      body: { countryCode, vatNumber },
    });
    if (error) throw error;
    return data;
  }

  // ===== VÉRIFICATION TVA EN DIRECT DANS LES MODALS (aperçu avant sauvegarde) =====
  async function checkTvaLive(numeroTva) {
    const tvaClean = (numeroTva || '').replace(/\s|\./g, '').toUpperCase();
    if (tvaClean.length < 9) { showToast('⚠️ Numéro TVA trop court'); return; }
    const countryCode = tvaClean.slice(0, 2);
    const vatNumber = tvaClean.slice(2);
    setTvaCheck({ loading: true });
    try {
      const json = await callViesRelay(countryCode, vatNumber);
      const statut = json.valid ? 'valide' : 'invalide';
      const nom = json.name && json.name !== '---' ? json.name : null;
      const adresse = json.address && json.address !== '---' ? json.address : null;
      setTvaCheck({ loading: false, statut, nom, adresse, pays: countryCode });
      if (statut === 'valide' && nom) {
        // Auto-remplir le nom si le formulaire est vide
        if (!donneurForm.nom) setDonneurForm((prev) => ({ ...prev, nom }));
      }
    } catch (e) {
      console.error(e);
      setTvaCheck({ loading: false, statut: 'erreur' });
      showToast('⚠️ Erreur de connexion au registre VIES (Union Européenne)');
    }
  }

  // ===== VÉRIFICATION TVA INTRACOMMUNAUTAIRE (API VIES — Union Européenne) =====
  async function verifierTVA(numeroTva, entityId, entityType) {
    const tvaClean = (numeroTva || '').replace(/\s|\./g, '').toUpperCase();
    if (tvaClean.length < 9) { showToast('⚠️ Numéro TVA trop court'); return; }
    const countryCode = tvaClean.slice(0, 2);
    const vatNumber = tvaClean.slice(2);
    try {
      const json = await callViesRelay(countryCode, vatNumber);
      const statut = json.valid ? 'valide' : 'invalide';
      const nomOfficiel = json.name && json.name !== '---' ? json.name : null;
      const table = entityType === 'st' ? 'sous_traitants' : 'donneurs_ordre';
      const updateData = { tva_verifie: true, tva_statut: statut, tva_pays: countryCode, tva_nom_officiel: nomOfficiel };
      const { error } = await supabase.from(table).update(updateData).eq('id', entityId);
      if (error) throw error;
      await loadAll();
      if (statut === 'valide') showToast(`✅ TVA valide${nomOfficiel ? ` — ${nomOfficiel}` : ''}`);
      else showToast('⚠️ Numéro TVA invalide selon le registre européen VIES');
    } catch (e) {
      console.error(e);
      showToast('⚠️ Erreur de connexion au registre VIES (Union Européenne)');
    }
  }

  // ===== SCORE DE SOLVABILITÉ (clients uniquement) =====
  // Calcul basé sur les données officielles de l'API gouvernementale française.
  // Score de 0 à 100 — plus le score est élevé, plus l'entreprise est crédible.
  async function calculerSolvabilite(siret, clientId) {
    const siretClean = (siret || '').replace(/\s/g, '');
    if (!siretClean) { showToast('⚠️ Renseigne le SIRET du client d\'abord'); return; }
    try {
      const res = await fetch(`https://recherche-entreprises.api.gouv.fr/search?q=${siretClean}&per_page=1`);
      const json = await res.json();
      const result = json.results && json.results[0];
      if (!result) { showToast('⚠️ Entreprise introuvable'); return; }

      const siege = result.siege || {};
      const etatAdmin = siege.etat_administratif || result.etat_administratif;
      const dateCreation = result.date_creation;
      const capital = result.capital || 0;
      const trancheEffectifs = result.tranche_effectif_salarie || '00';

      let score = 0;
      const details = {};

      // Entreprise trouvée dans la base officielle
      score += 15;
      details.trouvee = true;

      // Statut actif
      if (etatAdmin === 'A') { score += 25; details.active = true; }
      else { details.active = false; }

      // Ancienneté
      if (dateCreation) {
        const annees = Math.floor((Date.now() - new Date(dateCreation)) / (1000 * 60 * 60 * 24 * 365));
        details.anciennete = annees;
        if (annees >= 10) score += 25;
        else if (annees >= 5) score += 15;
        else if (annees >= 2) score += 8;
        else if (annees >= 1) score += 3;
      }

      // Capital social
      if (capital > 0) {
        details.capital = capital;
        if (capital >= 100000) score += 20;
        else if (capital >= 50000) score += 15;
        else if (capital >= 10000) score += 10;
        else score += 5;
      }

      // Effectifs
      const eff = parseInt(trancheEffectifs) || 0;
      details.trancheEffectifs = trancheEffectifs;
      if (eff >= 50) score += 15;
      else if (eff >= 10) score += 10;
      else if (eff >= 3) score += 5;
      else score += 2;

      score = Math.min(100, score);
      details.nomOfficiel = result.nom_complet || null;
      details.formeJuridique = result.nature_juridique || null;

      const { error } = await supabase.from('donneurs_ordre').update({
        score_solvabilite: score,
        score_details: details,
        score_date: new Date().toISOString(),
      }).eq('id', clientId);
      if (error) throw error;
      await loadAll();
      const mention = score >= 70 ? '🟢 Solide' : score >= 45 ? '🟡 Correct' : '🔴 Risqué';
      showToast(`✅ Score calculé : ${score}/100 — ${mention}`);
    } catch (e) {
      console.error(e);
      showToast('⚠️ Erreur lors du calcul de solvabilité');
    }
  }

  // ===== SCORE DE SOLVABILITÉ VIA TVA (entreprises étrangères sans SIRET) =====
  // Le registre VIES ne fournit pas de données financières (capital, ancienneté,
  // effectifs) comme l'API française — seulement la validité du numéro et le
  // nom associé. Le score reflète donc l'existence confirmée de l'entreprise,
  // pas sa solidité financière réelle. Ce mode complète le score par SIRET
  // pour les clients européens qui n'ont pas de SIRET français.
  async function calculerSolvabiliteTva(numeroTva, clientId) {
    const tvaClean = (numeroTva || '').replace(/\s|\./g, '').toUpperCase();
    if (tvaClean.length < 9) { showToast('⚠️ Numéro TVA trop court'); return; }
    const countryCode = tvaClean.slice(0, 2);
    const vatNumber = tvaClean.slice(2);
    try {
      const json = await callViesRelay(countryCode, vatNumber);
      const valide = !!json.valid;
      const score = valide ? 55 : 10;
      const details = {
        source: 'tva',
        trouvee: valide,
        nomOfficiel: json.name && json.name !== '---' ? json.name : null,
        pays: countryCode,
      };
      const { error } = await supabase.from('donneurs_ordre').update({
        score_solvabilite: score,
        score_details: details,
        score_date: new Date().toISOString(),
      }).eq('id', clientId);
      if (error) throw error;
      await loadAll();
      const mention = valide ? '🟡 Existence confirmée (TVA)' : '🔴 TVA invalide';
      showToast(`✅ Score : ${score}/100 — ${mention}`);
    } catch (e) {
      console.error(e);
      showToast('⚠️ Erreur lors du calcul de solvabilité par TVA');
    }
  }

  // ===== VÉRIFICATION SIRET (API officielle gouv.fr — recherche-entreprises) =====
  async function verifierSiret(siret, stId) {
    const siretClean = (siret || '').replace(/\s/g, '');
    if (siretClean.length < 9) {
      showToast('⚠️ SIRET/SIREN trop court pour vérification');
      return;
    }
    try {
      const res = await fetch(`https://recherche-entreprises.api.gouv.fr/search?q=${siretClean}&per_page=1`);
      const json = await res.json();
      const result = json.results && json.results[0];

      let statut = 'introuvable';
      let nomOfficiel = null;

      if (result) {
        const etatAdmin = result.siege?.etat_administratif || result.etat_administratif;
        statut = etatAdmin === 'A' ? 'actif' : 'cesse';
        nomOfficiel = result.nom_complet || result.nom_raison_sociale || null;
      }

      const { error } = await supabase.from('sous_traitants').update({
        siret_verifie: true,
        siret_statut: statut,
        siret_nom_officiel: nomOfficiel,
        siret_derniere_verif: new Date().toISOString(),
      }).eq('id', stId);

      if (error) throw error;

      await loadAll();
      if (statut === 'actif') {
        showToast(`✅ SIRET vérifié : ${nomOfficiel}`);
      } else if (statut === 'cesse') {
        showToast(`⚠️ Entreprise radiée/cessée : ${nomOfficiel}`);
      } else {
        showToast('⚠️ SIRET introuvable dans la base officielle');
      }
    } catch (e) {
      console.error(e);
      showToast('⚠️ Erreur de connexion à la vérification SIRET');
    }
  }

  async function createDonneur() {
    if (!donneurForm.nom) {
      showToast('⚠️ Le nom est requis');
      return;
    }
    const { error } = await supabase.from('donneurs_ordre').insert({
      nom: donneurForm.nom,
      type: donneurForm.type,
      delai_paiement: donneurForm.delai,
      statut: 'actif',
      numero_tva: donneurForm.numeroTva || null,
    });

    if (error) {
      showToast("⚠️ Erreur lors de l'ajout");
      console.error(error);
      return;
    }

    await loadAll();
    setDonneurModal(false); setTvaCheck(null);
    setDonneurForm({ nom: '', type: 'Commissionnaire de transport', delai: '', siret: '', numeroTva: '' });
    showToast(`✅ ${donneurForm.nom} ajouté — numéro client attribué automatiquement`);
  }

  async function deleteMission(ref) {
    const { error } = await supabase.from('missions').delete().eq('ref', ref);
    if (error) {
      showToast('⚠️ Erreur lors de la suppression');
      console.error(error);
      return;
    }
    await loadAll();
    showToast('Ordre de transport supprimé');
  }

  async function updateMissionStatut(ref, statut) {
    const { error } = await supabase.from('missions').update({ statut }).eq('ref', ref);
    if (error) { showToast('⚠️ Erreur mise à jour statut'); return; }
    await loadAll();
    const labels = { planifiee: 'Planifié', charge: 'Chargé 📦', en_cours: 'En cours 🚛', livree: 'Délivré ✓', annule: 'Annulé ✕', litige: 'Litige ⚠' };
    showToast(`✅ Statut : ${labels[statut] || statut}`);
  }
  async function nextFactureNumero(sens) {
    const prefix = sens === 'client' ? 'FA-CLIENT-' : 'FA-FOURNISSEUR-';
    const existing = factures.filter((f) => f.sens === sens);
    const maxNum = existing.reduce((max, f) => {
      const match = f.numero.match(/(\d+)$/);
      const n = match ? parseInt(match[1], 10) : 0;
      return Math.max(max, n);
    }, 0);
    return `${prefix}${String(maxNum + 1).padStart(3, '0')}`;
  }

  function openFactureModalFromMission(mission) {
    const sens = mission.flux === 'in' ? 'client' : 'fournisseur';
    const montant = mission.flux === 'in' ? mission.vendu : mission.paye;
    const today = new Date().toISOString().slice(0, 10);
    const echeance = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
    // Pour un ordre reçu (client), la facturation doit référencer le numéro
    // d'ordre fourni par le client — c'est ce numéro qu'il reconnaîtra —
    // tout en gardant une trace de la référence interne IBK (ex : B002).
    const missionRef = mission.flux === 'in' && mission.refClient
      ? `${mission.refClient} (réf. interne ${mission.ref})`
      : mission.ref;
    setFactureForm({
      sens,
      missionRef,
      tiers: mission.partenaire,
      date: today,
      echeance,
      montantHT: String(montant || ''),
      notes: '',
    });
    setFactureModal(true);
  }

  function openFactureModalBlank() {
    const today = new Date().toISOString().slice(0, 10);
    const echeance = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString().slice(0, 10);
    setFactureForm({ sens: 'client', missionRef: '', tiers: '', date: today, echeance, montantHT: '', notes: '' });
    setFactureModal(true);
  }

  async function createFacture() {
    if (!factureForm.tiers || !factureForm.montantHT) {
      showToast('⚠️ Renseigne au moins le tiers et le montant');
      return;
    }
    const numero = await nextFactureNumero(factureForm.sens);
    const { error } = await supabase.from('factures').insert({
      numero,
      sens: factureForm.sens,
      mission_ref: factureForm.missionRef || null,
      tiers: factureForm.tiers,
      date_facture: factureForm.date || new Date().toISOString().slice(0, 10),
      date_echeance: factureForm.echeance || null,
      montant_ht: parseFloat(factureForm.montantHT) || 0,
      // Affrètement intracommunautaire : aucune TVA appliquée (autoliquidation).
      taux_tva: 0,
      statut: 'emise',
      notes: factureForm.notes || null,
    });

    if (error) {
      showToast('⚠️ Erreur lors de la création de la facture');
      console.error(error);
      return;
    }

    await loadAll();
    setFactureModal(false);
    showToast(`✅ Facture ${numero} créée`);
  }

  async function updateFactureStatut(id, statut) {
    const { error } = await supabase.from('factures').update({ statut }).eq('id', id);
    if (error) {
      showToast('⚠️ Erreur lors de la mise à jour');
      console.error(error);
      return;
    }
    await loadAll();
    showToast(statut === 'payee' ? '✅ Facture marquée comme payée' : 'Statut mis à jour');
  }

  async function deleteFacture(id) {
    const { error } = await supabase.from('factures').delete().eq('id', id);
    if (error) {
      showToast('⚠️ Erreur lors de la suppression');
      console.error(error);
      return;
    }
    await loadAll();
    showToast('Facture supprimée');
  }

  async function updateDatePaiement(id, date) {
    const { error } = await supabase.from('factures').update({ date_paiement: date || null }).eq('id', id);
    if (error) { showToast('⚠️ Erreur mise à jour date'); return; }
    await loadAll();
    showToast('✅ Date de paiement enregistrée');
  }

  async function validerComptable(id) {
    const { error } = await supabase.from('factures').update({ valide_comptable: true, statut: 'payee' }).eq('id', id);
    if (error) { showToast('⚠️ Erreur validation'); return; }
    await loadAll();
    showToast('✅ Facture validée par la comptable');
  }
  function openAttestationModal(a) {
    setAttestationForm({
      id: a.id,
      licenceDateDebut: a.licenceDateDebut || '',
      licenceDate: a.licenceDate || '',
      assuranceDateDebut: a.assuranceDateDebut || '',
      assuranceDate: a.assuranceDate || '',
      kbisDateDebut: a.kbisDateDebut || '',
      kbisDate: a.kbisDate || '',
      chauffeurNom: a.chauffeurNom || '',
      chauffeurTel: a.chauffeurTel || '',
    });
    setAttestationModal(true);
  }

  async function saveAttestation() {
    const { error } = await supabase.from('attestations').update({
      licence_date_debut: attestationForm.licenceDateDebut || null,
      licence_date: attestationForm.licenceDate || null,
      assurance_date_debut: attestationForm.assuranceDateDebut || null,
      assurance_date: attestationForm.assuranceDate || null,
      kbis_date_debut: attestationForm.kbisDateDebut || null,
      kbis_date: attestationForm.kbisDate || null,
      chauffeur_nom: attestationForm.chauffeurNom || null,
      chauffeur_tel: attestationForm.chauffeurTel || null,
    }).eq('id', attestationForm.id);

    if (error) {
      showToast('⚠️ Erreur lors de la mise à jour');
      console.error(error);
      return;
    }
    await loadAll();
    setAttestationModal(false);
    showToast('✅ Attestation mise à jour');
  }

  function copyAttestationLink(a) {
    if (!a.publicToken) {
      showToast('⚠️ Lien indisponible pour ce sous-traitant');
      return;
    }
    const url = `${window.location.origin}${window.location.pathname}?token=${a.publicToken}`;
    navigator.clipboard.writeText(url);
    showToast('✅ Lien copié — envoie-le au sous-traitant');
  }

  async function viewDocument(path) {
    if (!path) {
      showToast('⚠️ Aucun document déposé pour le moment');
      return;
    }
    const { data, error } = await supabase.storage.from('documents-st').createSignedUrl(path, 3600);
    if (error || !data) {
      showToast('⚠️ Impossible d\'ouvrir le document');
      console.error(error);
      return;
    }
    window.open(data.signedUrl, '_blank');
  }

  function exportFacturePdf(facture) {
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

    // En-tête clair avec liseré doré en bas
    doc.setFillColor(250, 250, 251);
    doc.rect(0, 0, 210, 32, 'F');
    doc.setDrawColor(201, 168, 76);
    doc.setLineWidth(0.8);
    doc.line(0, 32, 210, 32);
    // Logo dans un cadre fin en haut à gauche
    try {
      doc.setDrawColor(225, 225, 228);
      doc.setLineWidth(0.3);
      doc.roundedRect(12, 5, 22, 22, 2, 2, 'S');
      doc.addImage(LOGO_IBK_BASE64, 'JPEG', 13.5, 6.5, 19, 19);
    } catch (e) {
      // Si l'image ne charge pas, on continue sans bloquer la génération du PDF
    }
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(13);
    doc.setTextColor(13, 27, 42);
    doc.text(SOCIETE.nom, 40, 14);
    doc.setFontSize(8);
    doc.setTextColor(110, 110, 115);
    doc.text(`${SOCIETE.adresse} — ${SOCIETE.ville}`, 40, 20);
    doc.text(`SIRET ${SOCIETE.siret}`, 40, 25);
    doc.setFontSize(15);
    doc.setTextColor(13, 27, 42);
    doc.text('FACTURE', 150, 14);
    doc.setFontSize(9);
    doc.setTextColor(110, 110, 115);
    doc.text('N° ' + facture.numero, 150, 21);
    doc.text('Date : ' + fmtDate(facture.date), 150, 26);

    let y = 44;
    doc.setDrawColor(200, 200, 200);
    doc.rect(14, y, 182, 26);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.5);
    doc.setTextColor(120, 120, 120);
    doc.text(facture.sens === 'client' ? 'FACTURÉ À' : 'FOURNISSEUR / SOUS-TRAITANT', 17, y + 6);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(11);
    doc.setTextColor(20, 20, 20);
    doc.text(facture.tiers, 17, y + 14);
    if (facture.missionRef) {
      doc.setFontSize(8.5);
      doc.setTextColor(100, 100, 100);
      doc.text('Mission associée : ' + facture.missionRef, 17, y + 21);
    }
    y += 32;

    doc.autoTable({
      startY: y,
      head: [['Description', 'Montant net à payer']],
      body: [[
        facture.sens === 'client'
          ? `Prestation de transport intracommunautaire — ${facture.missionRef || 'voir détail'}`
          : `Sous-traitance transport intracommunautaire — ${facture.missionRef || 'voir détail'}`,
        fmtEUR(facture.montantHT),
      ]],
      headStyles: { fillColor: [13, 27, 42], textColor: [201, 168, 76], fontStyle: 'bold', fontSize: 9 },
      bodyStyles: { fontSize: 9, cellPadding: 3 },
      margin: { left: 14, right: 14 },
      theme: 'grid',
    });

    let y2 = doc.lastAutoTable.finalY + 10;
    doc.setDrawColor(200, 200, 200);
    doc.rect(120, y2, 76, 20);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(11);
    doc.setTextColor(13, 27, 42);
    doc.text('Total à payer', 124, y2 + 8);
    doc.text(fmtEUR(facture.montantHT), 192, y2 + 8, { align: 'right' });
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(120, 120, 120);
    doc.text(doc.splitTextToSize(MENTION_TVA, 70), 124, y2 + 14);

    if (facture.echeance) {
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      doc.setTextColor(100, 100, 100);
      doc.text("Échéance de paiement : " + fmtDate(facture.echeance), 14, y2 + 7);
    }
    if (facture.notes) {
      doc.setFontSize(8.5);
      doc.text('Notes : ' + facture.notes, 14, y2 + 14, { maxWidth: 90 });
    }

    // Bloc coordonnées bancaires complètes (RIB)
    let y3 = y2 + 28;
    doc.setDrawColor(220, 220, 220);
    doc.setFillColor(248, 249, 251);
    doc.rect(14, y3, 182, 24, 'FD');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7.5);
    doc.setTextColor(120, 120, 120);
    doc.text('COORDONNÉES BANCAIRES', 18, y3 + 6);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(8.5);
    doc.setTextColor(40, 40, 40);
    doc.text(`Banque : ${SOCIETE.banque}`, 18, y3 + 12);
    doc.text(`IBAN : ${SOCIETE.iban}`, 18, y3 + 17.5);
    doc.text(`BIC : ${SOCIETE.bic}`, 120, y3 + 12);
    doc.text(`Titulaire : ${SOCIETE.nom}`, 120, y3 + 17.5);

    // Mention légale TVA (autoliquidation — transport intracommunautaire)
    let y4 = y3 + 30;
    doc.setFont('helvetica', 'italic');
    doc.setFontSize(7.5);
    doc.setTextColor(90, 90, 90);
    doc.text(doc.splitTextToSize(MENTION_TVA, 182), 14, y4);

    doc.setDrawColor(225, 225, 228);
    doc.setLineWidth(0.3);
    doc.line(0, 287, 210, 287);
    doc.setFontSize(7);
    doc.setTextColor(150, 150, 155);
    doc.text('IBK EURO AFRIQUE — Document généré par IBK TMS', 14, 293);

    doc.save(`${facture.numero}.pdf`);
    showToast('✅ Facture téléchargée en PDF');
  }


  // Charge une mission existante dans le formulaire "Ordre de transport"
  // pour en régénérer le PDF (mode réimpression : aucune nouvelle mission
  // n'est créée, otExistingRef indique qu'il s'agit d'une mission existante).
  function goToOT(ref) {
    const m = missions.find((x) => x.ref === ref);
    if (m) {
      setLv({
        ref: m.ref,
        date: m.date || '',
        dateLivraison: m.dateLivraison || '',
        affreteur: m.affreteur || '',
        affreteurEmail: m.affreteurEmail || '',
        affreteurTel: m.affreteurTel || '',
        transporteur: m.partenaire || '',
        transporteurAdresse: m.transporteurAdresse || '',
        transporteurContact: m.transporteurContact || '',
        transporteurContactEmail: m.transporteurContactEmail || '',
        transporteurContactTel: m.transporteurContactTel || '',
        plaque: m.plaque || '',
        lieuCharge: m.depart || '',
        cpCharge: m.cpCharge || '',
        paysCharge: m.paysCharge || 'France',
        lieuLivre: m.dest || '',
        cpLivre: m.cpLivre || '',
        paysLivre: m.paysLivre || 'France',
        clientChargement: m.contactChargement || '',
        heureChargementDebut: m.heureChargement || '',
        heureChargementFin: m.heureChargementFin || '',
        clientLivraison: m.contactLivraison || '',
        heureLivraisonDebut: m.heureLivraison || '',
        heureLivraisonFin: m.heureLivraisonFin || '',
        marchandise: m.marchandise || '',
        poids: m.poids || '',
        ldm: m.ldm || '',
        exchangePalettes: !!m.exchangePalettes,
        prixClient: String(m.vendu || ''),
        prix: String(m.paye || ''),
        delaiPaiement: m.delaiPaiement || '30 days end of month',
      });
      setOtExistingRef(ref);
    }
    setView('ordre-transport');
  }

  // Unique point de création d'une mission (ordre transporteur) : valide le
  // formulaire, crée la mission en base si elle n'existe pas encore
  // (otExistingRef vide), puis génère le PDF. Si on est en train de
  // régénérer le PDF d'une mission déjà créée, on saute directement à la
  // génération sans réinsertion.
  async function handleGenerateOT() {
    if (!lv.transporteur || !lv.lieuCharge || !lv.lieuLivre) {
      showToast('⚠️ Renseigne au moins le transporteur, le lieu de chargement et le lieu de livraison');
      return;
    }

    if (otExistingRef) {
      // Mission déjà créée : on ne fait que régénérer le PDF.
      exportOTPdf(otExistingRef);
      return;
    }

    const ref = `OT-${Date.now().toString().slice(-8)}`;
    const { error } = await supabase.from('missions').insert({
      ref,
      flux: 'out',
      source: 'Ordre de transport',
      partenaire: lv.transporteur,
      depart: lv.lieuCharge,
      destination: lv.lieuLivre,
      date_mission: lv.date || new Date().toISOString().slice(0, 10),
      date_livraison: lv.dateLivraison || null,
      vendu: parseFloat(lv.prixClient) || 0,
      paye: parseFloat(lv.prix) || 0,
      statut: 'planifiee',
      affreteur: lv.affreteur || null,
      affreteur_email: lv.affreteurEmail || null,
      affreteur_tel: lv.affreteurTel || null,
      transporteur_adresse: lv.transporteurAdresse || null,
      transporteur_contact: lv.transporteurContact || null,
      transporteur_contact_email: lv.transporteurContactEmail || null,
      transporteur_contact_tel: lv.transporteurContactTel || null,
      plaque: lv.plaque || null,
      contact_chargement: lv.clientChargement || null,
      heure_chargement: lv.heureChargementDebut || null,
      heure_chargement_fin: lv.heureChargementFin || null,
      contact_livraison: lv.clientLivraison || null,
      heure_livraison: lv.heureLivraisonDebut || null,
      heure_livraison_fin: lv.heureLivraisonFin || null,
      cp_charge: lv.cpCharge || null,
      pays_charge: lv.paysCharge || null,
      cp_livre: lv.cpLivre || null,
      pays_livre: lv.paysLivre || null,
      marchandise: lv.marchandise || null,
      poids: lv.poids || null,
      ldm: lv.ldm || null,
      exchange_palettes: !!lv.exchangePalettes,
      delai_paiement: lv.delaiPaiement || null,
    });

    if (error) {
      showToast('⚠️ Erreur lors de la création de la mission');
      console.error(error);
      return;
    }

    setLv((prev) => ({ ...prev, ref }));
    setOtExistingRef(ref);
    await loadAll();
    showToast(`✅ Mission ${ref} créée`);
    exportOTPdf(ref);
  }

  // cotation calc
  // Tarifs indicatifs €/km par type de camion — à ajuster selon ta grille
  // réelle si besoin (aucun tarif n'était précisé pour ces nouvelles
  // catégories dans le cahier des charges).
  const TARIFS_CAMION = {
    'Tautliner': 1.15,
    'Mega Tautliner': 1.45,
    'Box': 1.10,
    'Frigo': 1.30,
  };
  // Marge standard appliquée automatiquement (remplace l'ancien champ
  // "Marge cible" saisi manuellement, supprimé du formulaire).
  const MARGE_STANDARD_PCT = 20;

  const km = parseFloat(cotKm) || 0;
  const tarif = TARIFS_CAMION[cotTypeCamion] || 0;
  const peage = parseFloat(cotPeage) || 0;
  const cotBase = km * tarif;
  const cotRevient = cotBase + peage;
  const cotFinal = cotRevient * (1 + MARGE_STANDARD_PCT / 100);

  function useCotationInMission() {
    resetLv();
    setLv((prev) => ({
      ...prev,
      prixClient: cotFinal.toFixed(2),
      cpCharge: cotCpDepart,
      paysCharge: cotPaysDepart,
      cpLivre: cotCpArrivee,
      paysLivre: cotPaysArrivee,
    }));
    setView('ordre-transport');
    showToast('Prix de vente conseillé appliqué — complète le formulaire Ordre de transport');
  }

  function exportExcel() {
    const wb = XLSX.utils.book_new();
    const data = [['Référence', 'Flux', 'Source', 'Partenaire', 'Départ', 'Destination', 'Date', 'Prix vendu (€)', 'Prix payé (€)', 'Marge (€)', 'Marge %', 'Statut']];
    missions.forEach((m) => {
      const marge = m.flux === 'out' ? m.vendu - m.paye : '';
      const pct = m.flux === 'out' && m.vendu ? +(((m.vendu - m.paye) / m.vendu) * 100).toFixed(1) : '';
      data.push([m.ref, m.flux === 'out' ? 'Donné' : 'Reçu', m.source, m.partenaire, m.depart, m.dest, fmtDate(m.date), m.vendu, m.flux === 'out' ? m.paye : '', marge, pct, STATUS_MAP[m.statut]?.label || m.statut]);
    });
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [16, 10, 18, 20, 14, 14, 12, 14, 14, 12, 10, 12].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, 'Missions');

    const stData = [['Nom', 'SIRET', 'Contact', 'Téléphone', 'Flotte', 'Zone']];
    sousTraitants.forEach((s) => stData.push([s.nom, s.siret, s.contact, s.tel, s.flotte, s.zone]));
    const ws2 = XLSX.utils.aoa_to_sheet(stData);
    ws2['!cols'] = [26, 18, 16, 14, 20, 18].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws2, 'Sous-traitants');

    XLSX.writeFile(wb, 'IBK_TMS_Export.xlsx');
    showToast('✅ Export Excel téléchargé');
  }

  function exportOTPdf(refOverride) {
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    // Référence effective à utiliser dans tout le document : celle passée en
    // paramètre (juste après création de la mission, avant que le state React
    // n'ait eu le temps de se mettre à jour) ou, à défaut, celle du formulaire.
    const activeRef = refOverride || lv.ref;
    // Nombre de pages suivi dynamiquement : les CGT anglaises ci-dessous sont
    // volontairement reproduites à l'identique et peuvent nécessiter plus de
    // place qu'une simple page (le cahier des charges autorise une page
    // supplémentaire si besoin). Le total exact est corrigé en fin de
    // génération (voir la boucle finale plus bas), donc "0" est utilisé
    // comme valeur temporaire tant que le nombre final n'est pas connu.
    let pageNum = 1;

    // ===== FONCTION HELPER EN-TÊTE (réutilisée sur chaque page) =====
    function drawHeader(pageN, totalPages) {
      doc.setFillColor(13, 27, 42);
      doc.rect(0, 0, 210, 28, 'F');
      try { doc.addImage(LOGO_IBK_BASE64, 'JPEG', 8, 4, 20, 20); } catch (e) {}
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(11);
      doc.setTextColor(255, 255, 255);
      doc.text(SOCIETE.nom, 33, 13);
      doc.setFontSize(7.5);
      doc.setTextColor(180, 180, 180);
      doc.text(`${SOCIETE.adresse} — ${SOCIETE.ville} — SIRET ${SOCIETE.siret}`, 33, 20);
      // Titre à droite
      doc.setFontSize(13);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(212, 0, 110); // rose IBK
      doc.text('TRANSPORT ORDER', 210 - 14, 12, { align: 'right' });
      doc.setFontSize(8);
      doc.setTextColor(180, 180, 180);
      doc.text(`No. ${activeRef || '—'}  |  Date: ${lv.date ? fmtDate(lv.date) : '—'}  |  Page ${pageN}/${totalPages || '…'}`, 210 - 14, 20, { align: 'right' });
      // Liseré rose sous l'en-tête
      doc.setDrawColor(212, 0, 110);
      doc.setLineWidth(0.8);
      doc.line(0, 28, 210, 28);
    }

    function drawFooter() {
      doc.setDrawColor(200, 200, 200);
      doc.setLineWidth(0.3);
      doc.line(14, 286, 196, 286);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7);
      doc.setTextColor(150, 150, 155);
      doc.text('IBK EURO AFRIQUE — Intra-community transport — Confidential document generated by IBK TMS', 14, 291);
      doc.text(`${SOCIETE.adresse}, ${SOCIETE.ville}`, 210 - 14, 291, { align: 'right' });
    }

    function sectionTitle(text, y) {
      doc.setFillColor(212, 0, 110);
      doc.rect(14, y, 182, 6, 'F');
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7.5);
      doc.setTextColor(255, 255, 255);
      doc.text(text.toUpperCase(), 17, y + 4.3);
      return y + 8;
    }

    function infoBox(label, value, x, y, w, h) {
      doc.setDrawColor(220, 220, 220);
      doc.setLineWidth(0.3);
      doc.rect(x, y, w, h);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(7);
      doc.setTextColor(130, 130, 130);
      doc.text(label.toUpperCase(), x + 2.5, y + 4.5);
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      doc.setTextColor(13, 27, 42);
      doc.text(value || '—', x + 2.5, y + 10, { maxWidth: w - 5 });
    }

    const BOTTOM_LIMIT = 268;

    // Passe à une nouvelle page si besoin (contenu qui déborde), en
    // clôturant la page courante (pied de page) puis en ouvrant la
    // suivante avec un bandeau de continuité.
    function newCgtPage(label = 'General Terms and Conditions of Transport (continued)') {
      drawFooter();
      doc.addPage();
      pageNum++;
      drawHeader(pageNum, 0);
      let ny = 34;
      ny = sectionTitle(label, ny);
      return ny;
    }

    function ensureRoom(neededHeight, label = 'Transport Order (continued)') {
      if (y + neededHeight > BOTTOM_LIMIT) {
        y = newCgtPage(label);
      }
    }

    // ============================================================
    // PAGE 1 — INFORMATIONS PRINCIPALES
    // ============================================================
    drawHeader(1, 0);
    let y = 34;

    // Bandeau de référence
    doc.setFillColor(248, 248, 248);
    doc.rect(14, y, 182, 10, 'F');
    doc.setDrawColor(212, 0, 110);
    doc.setLineWidth(0.3);
    doc.rect(14, y, 182, 10);
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(9);
    doc.setTextColor(13, 27, 42);
    doc.text(`Ref.: ${activeRef || '—'}`, 18, y + 6.5);
    doc.text(`Pickup: ${lv.date ? fmtDate(lv.date) : '—'}`, 80, y + 6.5);
    doc.text(`Delivery: ${lv.dateLivraison ? fmtDate(lv.dateLivraison) : '—'}`, 145, y + 6.5);
    y += 14;

    // Affréteur IBK
    ensureRoom(24);
    y = sectionTitle('IBK Freight Broker (Negotiation)', y);
    infoBox('Name', lv.affreteur || 'IBK EURO AFRIQUE', 14, y, 60, 16);
    infoBox('Email', lv.affreteurEmail, 74, y, 62, 16);
    infoBox('Phone', lv.affreteurTel, 136, y, 60, 16);
    y += 20;

    // Transporteur (sous-traitant)
    ensureRoom(44);
    y = sectionTitle('Carrier Performing the Transport', y);
    infoBox('Carrier Name', lv.transporteur, 14, y, 88, 16);
    infoBox('Registration', lv.plaque, 108, y, 88, 16);
    y += 20;
    infoBox('Company Address', lv.transporteurAdresse, 14, y, 182, 14);
    y += 18;
    infoBox('Contact at Carrier', lv.transporteurContact, 14, y, 60, 16);
    infoBox('Contact Email', lv.transporteurContactEmail, 74, y, 62, 16);
    infoBox('Contact Phone', lv.transporteurContactTel, 136, y, 60, 16);
    y += 20;

    // Chargement
    ensureRoom(52);
    y = sectionTitle('Loading', y);
    infoBox('Loading Location', lv.lieuCharge, 14, y, 88, 16);
    infoBox('Postal Code', lv.cpCharge, 108, y, 40, 16);
    infoBox('Country', lv.paysCharge, 154, y, 42, 16);
    y += 20;
    infoBox('Loading Client', lv.clientChargement, 14, y, 88, 14);
    infoBox('Start Time', lv.heureChargementDebut, 108, y, 40, 14);
    infoBox('End Time', lv.heureChargementFin, 154, y, 42, 14);
    y += 18;

    // Livraison
    ensureRoom(52);
    y = sectionTitle('Delivery', y);
    infoBox('Delivery Location', lv.lieuLivre, 14, y, 88, 16);
    infoBox('Postal Code', lv.cpLivre, 108, y, 40, 16);
    infoBox('Country', lv.paysLivre, 154, y, 42, 16);
    y += 20;
    infoBox('Delivery Client', lv.clientLivraison, 14, y, 88, 14);
    infoBox('Start Time', lv.heureLivraisonDebut, 108, y, 40, 14);
    infoBox('End Time', lv.heureLivraisonFin, 154, y, 42, 14);
    y += 18;

    // Marchandise
    ensureRoom(40);
    y = sectionTitle('Goods Transported', y);
    infoBox('Nature of Goods', lv.marchandise, 14, y, 88, 16);
    infoBox('Estimated Weight', lv.poids, 108, y, 40, 16);
    infoBox('LDM (Linear Meters)', lv.ldm, 154, y, 42, 16);
    y += 20;
    infoBox('Pallet Exchange', lv.exchangePalettes ? 'Yes' : 'No', 14, y, 88, 14);
    y += 18;

    // Prix & délai
    ensureRoom(30);
    y = sectionTitle('Price & Payment', y);
    infoBox('Agreed Payment Terms', lv.delaiPaiement || '30 days end of month', 14, y, 128, 16);
    doc.setFillColor(13, 27, 42);
    doc.rect(146, y, 50, 16, 'F');
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(7);
    doc.setTextColor(212, 0, 110);
    doc.text('NET PRICE', 149, y + 5.5);
    doc.setFontSize(12);
    doc.setTextColor(255, 255, 255);
    doc.text(`${(parseFloat(lv.prix) || 0).toFixed(2)} €`, 149, y + 13);
    y += 20;

    // Mention légale TVA (autoliquidation) — reste visible même après la
    // suppression de la page de récapitulatif financier.
    doc.setFont('helvetica', 'italic');
    doc.setFontSize(6.5);
    doc.setTextColor(120, 120, 120);
    doc.text(doc.splitTextToSize(MENTION_TVA, 182), 14, y);
    y += 8;

    drawFooter();

    // ============================================================
    // PAGE(S) SUIVANTE(S) — GENERAL TERMS AND CONDITIONS OF TRANSPORT
    // Reproduites à l'identique (texte fourni), en anglais, structure et
    // ordre strictement respectés. S'étale automatiquement sur plusieurs
    // pages si nécessaire (le contenu est plus long qu'une page unique).
    // ============================================================
    doc.addPage();
    pageNum++;
    drawHeader(pageNum, 0);
    y = 34;
    y = sectionTitle('General Terms and Conditions of Transport', y);

    function cgtArticleTitle(text) {
      if (y + 8 > BOTTOM_LIMIT) y = newCgtPage();
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(9);
      doc.setTextColor(13, 27, 42);
      doc.text(text, 14, y);
      y += 6;
    }

    function cgtParagraph(text) {
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(7.5);
      const lines = doc.splitTextToSize(text, 182);
      const blockHeight = lines.length * 3.7 + 3;
      if (y + blockHeight > BOTTOM_LIMIT) y = newCgtPage();
      doc.setTextColor(60, 60, 60);
      doc.text(lines, 14, y);
      y += blockHeight;
    }

    function cgtBullets(items) {
      items.forEach((item) => {
        doc.setFont('helvetica', 'normal');
        doc.setFontSize(7.5);
        const lines = doc.splitTextToSize(`•  ${item}`, 176);
        const blockHeight = lines.length * 3.7 + 1;
        if (y + blockHeight > BOTTOM_LIMIT) y = newCgtPage();
        doc.setTextColor(60, 60, 60);
        doc.text(lines, 18, y);
        y += blockHeight;
      });
      y += 2;
    }

    // ---- Article 1 ----
    cgtArticleTitle('1. Acceptance of the Transport Order');
    cgtParagraph('The present Transport Order shall become legally binding upon the Service provider through any of the following:');
    cgtBullets([
      'any price agreement and Order sended to the services provider (carriers, brokers, freight forwarder; etc…) without confirmation (automatically accepted)',
      'Written confirmation by email, SMS, WhatsApp, or any electronic communication;',
      'Commencement of the transport services.',
    ]);
    cgtParagraph('These Terms and Conditions shall prevail over any terms and conditions issued by the Carrier unless expressly agreed otherwise in writing by the Company.');

    // ---- Article 2 ----
    cgtArticleTitle("2. Carrier's Responsibility");
    cgtParagraph('The Carrier shall bear full responsibility for the cargo from the time of collection until final delivery.');
    cgtParagraph('The Carrier shall be fully liable for any loss, theft, shortage, deterioration or damage occurring during this period, except where such loss is exclusively caused by duly proven force majeure.');

    // ---- Article 3 ----
    cgtArticleTitle('3. Theft, Loss and Shortages');
    cgtParagraph('The Carrier shall fully indemnify the Company and/or its Client for:');
    cgtBullets(['Loss of goods;', 'Theft;', 'Partial shortages;', 'Misdelivery;', 'Unauthorized delivery.']);
    cgtParagraph('Compensation shall include, without limitation:');
    cgtBullets(['Full value of the goods;', 'Freight charges;', 'Recovery expenses;', 'Legal costs;', 'Administrative costs;', 'Any direct or indirect financial losses.']);
    cgtParagraph('The Carrier remains fully liable for all acts or omissions committed by its employees, drivers, subcontractors, agents or any third party acting on its behalf.');

    // ---- Article 4 ----
    cgtArticleTitle('4. Insurance');
    cgtParagraph('The Carrier warrants that valid insurance policies remain in force throughout the transport operation, including:');
    cgtBullets(['Cargo Insurance;', 'Goods in Transit Insurance;', 'Theft Coverage;', 'Public Liability Insurance.']);
    cgtParagraph('Failure to maintain valid insurance shall entitle the Company to immediately cancel the transport order without compensation.');

    // ---- Article 5 ----
    cgtArticleTitle('5. Cargo Damage');
    cgtParagraph("Any damage identified upon delivery shall be presumed to have occurred while the goods were under the Carrier's custody unless the Carrier proves otherwise.");
    cgtParagraph('The Carrier shall bear all costs relating to:');
    cgtBullets(['Repair;', 'Replacement;', 'Expert assessments;', 'Disposal;', 'Additional transport;', 'Client compensation.']);

    // ---- Article 6 ----
    cgtArticleTitle('6. Delivery Deadlines');
    cgtParagraph('Delivery deadlines are mandatory.');
    cgtParagraph('Any delay attributable to the Carrier may result in financial penalties which the Company may deduct directly from any amount due to the Carrier.');

    // ---- Article 7 ----
    cgtArticleTitle('7. Cancellation by the Carrier');
    cgtParagraph('If the Carrier cancels the transport less than twenty-four (24) hours before loading, the Carrier shall be fully liable for:');
    cgtBullets(['Replacement carrier costs;', 'Client penalties;', 'Operational losses;', 'Administrative expenses.']);
    cgtParagraph('Additionally:');
    cgtBullets([
      'Any cancellation made more than two (2) hours after receipt of the Transport Order shall incur a cancellation fee ranging from EUR 50 to EUR 100, depending on the operational costs incurred.',
      'Any cancellation made on the scheduled loading day, whether several hours or only a few minutes before loading, shall incur cancellation charges ranging from EUR 200 to EUR 500, without prejudice to any additional damages suffered by the Company.',
    ]);

    drawFooter();

    // ============================================================
    // PASSE FINALE : maintenant que le nombre total de pages est connu,
    // on revient sur chaque page pour corriger l'indicateur "Page X/Y"
    // (jusqu'ici affiché avec un total provisoire "…").
    // ============================================================
    const totalPages = pageNum;
    for (let p = 1; p <= totalPages; p++) {
      doc.setPage(p);
      drawHeader(p, totalPages);
    }

    doc.save(`Ordre_de_transport_${activeRef || 'IBK'}.pdf`);
    showToast(`✅ Ordre de transport PDF (${totalPages} pages) téléchargé`);
  }

  // ===== STYLES (Tailwind utility classes throughout) =====
  if (loading) {
    return (
      <div className="min-h-screen bg-[#f5f5f5] flex items-center justify-center text-[#d4006e] font-semibold">
        Chargement des données…
      </div>
    );
  }

  const NAV_TITLES = {
    dashboard: 'Tableau de bord — IBK TMS',
    missions: "Missions d'affrètement",
    marges: 'Analyse des marges',
    affectation: 'Affectation — Tableau journalier',
    'historique-axes': 'Historique des axes',
    roles: 'Gestion de l\'équipe & des rôles',
    facturation: 'Facturation',
    'sous-traitants': 'Sous-traitants partenaires',
    attestations: 'Suivi des attestations',
    donneurs: "Clients",
    'demandes-devis': 'Demandes de devis (site vitrine)',
    cotation: 'Cotation rapide',
    'ordre-transport': 'Ordre de transport',
  };

  return (
    <div className="min-h-screen bg-[#f5f5f5] text-[#1c2733] flex font-sans" style={{ fontFamily: "'Inter', sans-serif" }}>
      <Toast message={toast} />

      {/* BANDE SUPÉRIEURE */}
      <div className="fixed top-0 left-0 right-0 h-12 bg-[#0d1b2a] z-50 flex items-center justify-between px-5 gap-6">
        <div className="flex items-center gap-2.5 flex-shrink-0">
          <img src={LOGO_IBK_BASE64} alt="IBK Euro Afrique" className="h-8 w-8 rounded-md object-cover" />
          <span className="text-white font-bold text-[15px] hidden sm:inline" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>IBK TMS</span>
        </div>
        <nav className="flex items-center gap-5 overflow-x-auto flex-1 min-w-0">
          {[
            ['dashboard', 'Tableau de bord'],
            ['missions', 'Ordres'],
            ['marges', 'Marges'],
            ['facturation', 'Facturation'],
            ['sous-traitants', 'Sous-traitants'],
            ['attestations', 'Vérification'],
            ['donneurs', 'Clients'],
            ['demandes-devis', 'Demandes'],
            ['affectation', 'Affectation'],
            ['cotation', 'Cotation'],
            ['ordre-transport', 'Ordre de transport'],
            ['historique-axes', 'Axes'],
            ['roles', 'Équipe'],
          ].filter(([key]) => {
            if (key === 'marges') return canSeeMarges;
            if (key === 'facturation') return canSeeFacturation;
            if (key === 'affectation') return canSeeAffectation;
            if (key === 'historique-axes') return canSeeAxes;
            if (key === 'sous-traitants') return canManageSousTraitants;
            if (key === 'donneurs') return canSeeClients;
            if (key === 'demandes-devis') return canSeeDemandes;
            if (key === 'roles') return isAdmin;
            return true;
          }).map(([key, label]) => (
            <span
              key={key}
              onClick={() => setView(key)}
              className={`whitespace-nowrap text-[12.5px] font-semibold cursor-pointer transition pb-1 border-b-2 flex items-center gap-1.5 ${
                view === key ? 'text-white border-white' : 'text-white/60 border-transparent hover:text-white/90'
              }`}
            >
              {label}
              {key === 'demandes-devis' && demandesDevis.filter((d) => d.statut === 'nouveau').length > 0 && (
                <span className="bg-[#d4006e] text-white text-[9px] font-bold rounded-full px-1.5 py-0.5 leading-none">
                  {demandesDevis.filter((d) => d.statut === 'nouveau').length}
                </span>
              )}
            </span>
          ))}
        </nav>

      </div>

      {/* SIDEBAR */}
      <aside className="w-56 bg-[#0d1b2a] text-white border-r border-[#d4006e]/30 shadow-[3px_0_20px_rgba(13,27,42,0.25)] flex flex-col fixed top-12 left-0 bottom-0 z-40">
        <div className="px-5 pt-6 pb-5 border-b border-[#d4006e]/20">
          <div className="text-[19px] font-bold text-white" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>IBK TMS</div>
          <div className="text-[10px] text-white/50 uppercase tracking-wider mt-0.5">Module persistant</div>
        </div>

        <div className="mx-5 mt-3.5 flex items-center justify-center gap-1.5 px-2 py-2 rounded-lg bg-white/10 border border-white/20 text-[10px] font-semibold">
          <span className="text-[#d4006e]">▲ Donné</span>
          <span className="text-white/50">⇄</span>
          <span className="text-white">▼ Reçu</span>
        </div>



        {isAdmin && (
          <div className="px-5 py-3 border-t border-[#d4006e]/20">
            <div className="text-[10px] uppercase tracking-wide text-white/50 mb-2">En ligne maintenant</div>
            {onlineUsers.length === 0 ? (
              <div className="text-[11px] text-white/40">Personne pour l'instant</div>
            ) : (
              onlineUsers.map((email) => (
                <div key={email} className="flex items-center gap-1.5 text-[11px] text-white/80 mb-1">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 flex-shrink-0"></span>
                  <span className="truncate">{email === userEmail ? `${email} (toi)` : email}</span>
                </div>
              ))
            )}
          </div>
        )}

        <div className="px-5 py-4 border-t border-[#d4006e]/20">
          <div className="flex items-center gap-2.5">
            <div className="w-8.5 h-8.5 rounded-full bg-[#d4006e] text-white flex items-center justify-center font-bold text-[13px]" style={{ fontFamily: "'Space Grotesk', sans-serif", width: 34, height: 34 }}>
              {(userEmail || '?').slice(0, 2).toUpperCase()}
            </div>
            <div className="overflow-hidden flex-1 min-w-0">
              <div className="text-[11px] font-semibold text-white break-all leading-snug" title={userEmail}>{userEmail}</div>
              <div className="text-[10px] text-white/50 mt-0.5">Données partagées équipe</div>
            </div>
          </div>
          <button
            onClick={onLogout}
            className="w-full mt-3 text-[11px] text-[#64748b] hover:text-[#d4006e] border border-[#d4006e]/20 rounded-md py-1.5 transition"
          >
            Se déconnecter
          </button>
        </div>
      </aside>

      {/* MAIN */}
      <main className="ml-56 mt-12 flex-1 flex flex-col min-h-screen">
        <div className="bg-[#ffffff] border-b border-[#d4006e]/20 shadow-[0_3px_16px_rgba(13,27,42,0.10)] px-7 py-3.5 flex items-center justify-between sticky top-0 z-30 flex-wrap gap-2">
          <div className="text-[18px] font-bold" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{NAV_TITLES[view]}</div>
          <div className="flex items-center gap-2.5">
            {(view === 'dashboard' || view === 'missions') && (
              <button onClick={exportExcel} className="px-3.5 py-2 rounded-md text-[13px] font-semibold bg-gray-100 border border-[#d4006e]/30 hover:border-[#d4006e] hover:text-[#d4006e] transition" style={{ background: '#f3f4f6' }}>⬇ Télécharger</button>
            )}
          </div>
        </div>

        <div className="p-7">
          {view === 'dashboard' && (
            <DashboardView
              totalOut={totalOut} totalIn={totalIn} totalMarge={totalMarge} margeAvgPct={margeAvgPct}
              missions={missions} factures={factures} sousTraitants={sousTraitants}
              outCount={outMissions.length} inCount={inMissions.length} isAdmin={isAdmin}
            />
          )}

          {view === 'missions' && (
            <MissionsView
              missions={filteredMissions} flowFilter={flowFilter} setFlowFilter={setFlowFilter}
              statusFilter={statusFilter} setStatusFilter={setStatusFilter}
              onDelete={deleteMission} onOT={goToOT} onFacturer={openFactureModalFromMission}
              onUpdateStatut={updateMissionStatut} onOpenDocs={(m) => setDocModal(m)} isAdmin={isAdmin}
            />
          )}

          {view === 'marges' && (isAdmin ? <MargesView outMissions={outMissions} totalMarge={totalMarge} margeAvgPct={margeAvgPct} onDelete={deleteMission} /> : <AccesRestreint />)}
          {view === 'affectation' && (isAdmin ? <AffectationView userEmail={userEmail} showToast={showToast} /> : <AccesRestreint />)}
          {view === 'historique-axes' && (canSeeAxes ? <HistoriqueAxesView /> : <AccesRestreint />)}
          {view === 'roles' && (isAdmin ? <RolesView showToast={showToast} /> : <AccesRestreint />)}

          {view === 'facturation' && (
            isAdmin ? (
              <FacturationView
                factures={factures}
                onAdd={openFactureModalBlank}
                onUpdateStatut={updateFactureStatut}
                onDelete={deleteFacture}
                onExportPdf={exportFacturePdf}
                onUpdateDatePaiement={updateDatePaiement}
                onValiderComptable={validerComptable}
              />
            ) : <AccesRestreint />
          )}

          {view === 'sous-traitants' && (
            <SousTraitantsView sousTraitants={sousTraitants} missions={missions} onAdd={() => { setStForm({ id: null, nom: '', siret: '', contact: '', tel: '', email: '', numeroTeleroute: '', numeroTimocom: '', flotte: '', zone: '', typeEntreprise: 'transporteur', numeroTva: '' }); setStSiretCheck(null); setTvaCheck(null); setStVerifMethod('tva'); setStModal(true); }} onVerifySiret={verifierSiret} onEdit={openEditSousTraitant} onDelete={deleteSousTraitant} onVerifierTVA={verifierTVA} />
          )}

          {view === 'attestations' && <AttestationsView attestations={attestations} sousTraitants={sousTraitants} stNameById={stNameById} onEdit={openAttestationModal} onViewDoc={viewDocument} onCopyLink={copyAttestationLink} />}

          {view === 'donneurs' && <DonneursView donneurs={donneurs} missions={missions} onAdd={() => { setDonneurVerifMethod('tva'); setDonneurModal(true); }} onVerifierTVA={verifierTVA} onSolvabilite={calculerSolvabilite} onSolvabiliteTva={calculerSolvabiliteTva} onDelete={deleteDonneur} />}
          {view === 'demandes-devis' && <DemandesDevisView demandes={demandesDevis} onUpdateStatut={updateDemandeStatut} onDelete={deleteDemande} onConvert={convertDemandeEnClient} />}

          {view === 'cotation' && (
            <CotationView
              cotPaysDepart={cotPaysDepart} setCotPaysDepart={setCotPaysDepart}
              cotCpDepart={cotCpDepart} setCotCpDepart={setCotCpDepart}
              cotPaysArrivee={cotPaysArrivee} setCotPaysArrivee={setCotPaysArrivee}
              cotCpArrivee={cotCpArrivee} setCotCpArrivee={setCotCpArrivee}
              cotKm={cotKm} setCotKm={setCotKm} cotTypeCamion={cotTypeCamion} setCotTypeCamion={setCotTypeCamion}
              cotPeage={cotPeage} setCotPeage={setCotPeage}
              cotBase={cotBase} cotRevient={cotRevient} cotFinal={cotFinal}
              margeStandardPct={MARGE_STANDARD_PCT}
              onUse={useCotationInMission}
            />
          )}

          {view === 'ordre-transport' && <OrdreTransportView lv={lv} setLv={setLv} sousTraitants={sousTraitants} otExistingRef={otExistingRef} onGenerate={handleGenerateOT} onReset={resetLv} />}
        </div>
      </main>

      {/* MODAL DOCUMENTS ORDRE REÇU (ordre client + CMR) */}
      {docModal && (
        <Modal title={`Documents — ${docModal.ref}`} onClose={() => setDocModal(null)} width={520}>
          <div className="text-[12px] text-[#64748b] mb-5">
            Dépose ici l'ordre de transport reçu du client (n° {docModal.refClient || '—'}) ainsi que la CMR associée. Ces documents servent à clôturer l'ordre et à établir la facture client.
          </div>

          <div className="mb-5 p-4 rounded-lg border border-[#d4006e]/20">
            <div className="text-[11px] uppercase tracking-wide text-[#64748b] mb-2 font-semibold">Ordre de transport du client</div>
            {docModal.ordreClientUrl ? (
              <div className="flex items-center justify-between gap-2 flex-wrap">
                <Badge cls="bg-emerald-500/15 text-emerald-700">✓ Document déposé</Badge>
                <button onClick={() => viewClientDocument(docModal.ordreClientUrl)} className="text-[11px] px-2.5 py-1.5 rounded border border-[#0d1b2a]/30 text-[#0d1b2a] hover:bg-gray-100">Voir le document</button>
              </div>
            ) : (
              <div className="flex items-center gap-2 flex-wrap">
                <input
                  type="file" accept=".pdf,image/*" className="form-input"
                  onChange={(e) => e.target.files[0] && uploadDocForMission(docModal.ref, e.target.files[0], 'ordre-client')}
                  disabled={docUploading}
                />
                {docUploading && <span className="text-[11px] text-[#64748b]">Téléversement en cours…</span>}
              </div>
            )}
          </div>

          <div className="mb-5 p-4 rounded-lg border border-[#d4006e]/20">
            <div className="text-[11px] uppercase tracking-wide text-[#64748b] mb-2 font-semibold">CMR</div>
            {docModal.cmrUrl ? (
              <div className="flex items-center justify-between gap-2 flex-wrap">
                <Badge cls="bg-emerald-500/15 text-emerald-700">✓ Document déposé</Badge>
                <button onClick={() => viewClientDocument(docModal.cmrUrl)} className="text-[11px] px-2.5 py-1.5 rounded border border-[#0d1b2a]/30 text-[#0d1b2a] hover:bg-gray-100">Voir le document</button>
              </div>
            ) : (
              <div className="flex items-center gap-2 flex-wrap">
                <input
                  type="file" accept=".pdf,image/*" className="form-input"
                  onChange={(e) => e.target.files[0] && uploadDocForMission(docModal.ref, e.target.files[0], 'cmr')}
                  disabled={docUploading}
                />
                {docUploading && <span className="text-[11px] text-[#64748b]">Téléversement en cours…</span>}
              </div>
            )}
          </div>

          <div className="flex justify-end gap-2.5 mt-2">
            <button onClick={() => setDocModal(null)} className="btn-ghost">Fermer</button>
          </div>
        </Modal>
      )}

      {/* MODAL SOUS-TRAITANT */}
      {stModal && (
        <Modal title={stForm.id ? 'Modifier le sous-traitant' : 'Ajouter un sous-traitant'} onClose={() => { setStModal(false); setStForm({ id: null, nom: '', siret: '', contact: '', tel: '', email: '', numeroTeleroute: '', numeroTimocom: '', flotte: '', zone: '', typeEntreprise: 'transporteur', numeroTva: '' }); setStSiretCheck(null); }} width={520}>
          <FormRow>
            <Field label="Raison sociale">
              <input className="form-input" value={stForm.nom} onChange={(e) => setStForm({ ...stForm, nom: e.target.value })} placeholder="Nom de l'entreprise" />
            </Field>
            <Field label="Type d'entreprise">
              <select className="form-input" value={stForm.typeEntreprise} onChange={(e) => setStForm({ ...stForm, typeEntreprise: e.target.value })}>
                <option value="transporteur">🚛 Transporteur</option>
                <option value="commissionnaire">🏢 Commissionnaire</option>
              </select>
            </Field>
          </FormRow>
          <div className="mb-1">
            <label className="text-[11px] uppercase tracking-wide text-[#64748b] mb-1.5 block">Méthode de vérification</label>
            <div className="flex bg-gray-100 rounded-lg p-1 mb-3" style={{ maxWidth: 340 }}>
              <button type="button" onClick={() => setStVerifMethod('tva')} className={`flex-1 py-1.5 rounded-md text-[12px] font-semibold transition ${stVerifMethod === 'tva' ? 'bg-[#0d1b2a] text-white' : 'text-[#64748b]'}`}>🇪🇺 Par TVA (recommandé)</button>
              <button type="button" onClick={() => setStVerifMethod('siret')} className={`flex-1 py-1.5 rounded-md text-[12px] font-semibold transition ${stVerifMethod === 'siret' ? 'bg-[#0d1b2a] text-white' : 'text-[#64748b]'}`}>🇫🇷 Par SIRET</button>
            </div>
          </div>

          {stVerifMethod === 'tva' ? (
            <>
              <Field label="N° TVA intracommunautaire">
                <input className="form-input" value={stForm.numeroTva} onChange={(e) => { setStForm({ ...stForm, numeroTva: e.target.value }); setTvaCheck(null); }} placeholder="ex : FR12345678901 / DE123456789 / BE0123456789" />
              </Field>
              <div className="mb-4 mt-2 flex items-center gap-2.5 flex-wrap">
                <button
                  type="button"
                  onClick={() => checkTvaLive(stForm.numeroTva)}
                  disabled={!stForm.numeroTva || tvaCheck?.loading}
                  className="text-[11px] px-3 py-1.5 rounded bg-[#0d1b2a] text-white disabled:opacity-40"
                >
                  {tvaCheck?.loading ? 'Vérification…' : 'Vérifier'}
                </button>
                {tvaCheck && !tvaCheck.loading && tvaCheck.statut === 'valide' && <Badge cls="bg-emerald-500/15 text-emerald-700">✓ TVA {tvaCheck.pays} valide{tvaCheck.nom ? ` — ${tvaCheck.nom}` : ''}</Badge>}
                {tvaCheck && !tvaCheck.loading && tvaCheck.statut === 'invalide' && <Badge cls="bg-red-500/15 text-red-700">✕ Numéro TVA invalide</Badge>}
                {tvaCheck && !tvaCheck.loading && tvaCheck.statut === 'erreur' && <Badge cls="bg-amber-400/15 text-amber-700">⚠ Erreur connexion VIES</Badge>}
              </div>
            </>
          ) : (
            <>
              <Field label="SIRET">
                <input className="form-input" value={stForm.siret} onChange={(e) => { setStForm({ ...stForm, siret: e.target.value }); setStSiretCheck(null); }} placeholder="000 000 000 00000" />
              </Field>
              <div className="mb-4 mt-2 flex items-center gap-2.5 flex-wrap">
                <button
                  type="button"
                  onClick={checkSiretLive}
                  disabled={!stForm.siret || stSiretCheck?.loading}
                  className="text-[11px] px-3 py-1.5 rounded bg-[#0d1b2a] text-white disabled:opacity-40"
                >
                  {stSiretCheck?.loading ? 'Vérification…' : 'Vérifier'}
                </button>
                {stSiretCheck && !stSiretCheck.loading && stSiretCheck.statut === 'actif' && (
                  <Badge cls="bg-emerald-500/15 text-emerald-700">✓ Actif — {stSiretCheck.nom}</Badge>
                )}
                {stSiretCheck && !stSiretCheck.loading && stSiretCheck.statut === 'actif' && stSiretCheck.adresse && (
                  <span className="text-[10px] text-[#64748b] basis-full">📍 {stSiretCheck.adresse}</span>
                )}
                {stSiretCheck && !stSiretCheck.loading && stSiretCheck.statut === 'actif' && stSiretCheck.dirigeant && (
                  <span className="text-[10px] text-[#64748b] basis-full">👤 Dirigeant : {stSiretCheck.dirigeant}</span>
                )}
                {stSiretCheck && !stSiretCheck.loading && stSiretCheck.statut === 'actif' && !stSiretCheck.dirigeant && (
                  <span className="text-[10px] text-[#64748b] basis-full">ℹ️ Nom du dirigeant non disponible pour cette entreprise (champ à compléter manuellement)</span>
                )}
                {stSiretCheck && !stSiretCheck.loading && stSiretCheck.statut === 'cesse' && (
                  <Badge cls="bg-red-500/15 text-red-700">⚠ Radiée/cessée — {stSiretCheck.nom}</Badge>
                )}
                {stSiretCheck && !stSiretCheck.loading && stSiretCheck.statut === 'introuvable' && (
                  <Badge cls="bg-red-500/15 text-red-700">✕ SIRET introuvable</Badge>
                )}
                {stSiretCheck && !stSiretCheck.loading && stSiretCheck.statut === 'erreur' && (
                  <Badge cls="bg-amber-400/15 text-amber-700">⚠ Erreur de connexion, réessaie</Badge>
                )}
              </div>
            </>
          )}


          <FormRow>
            <Field label="Contact">
              <input className="form-input" value={stForm.contact} onChange={(e) => setStForm({ ...stForm, contact: e.target.value })} placeholder="Nom du contact" />
            </Field>
            <Field label="Téléphone">
              <input className="form-input" value={stForm.tel} onChange={(e) => setStForm({ ...stForm, tel: e.target.value })} placeholder="06 00 00 00 00" />
            </Field>
          </FormRow>
          <Field label="Adresse e-mail">
            <input type="email" className="form-input" value={stForm.email} onChange={(e) => setStForm({ ...stForm, email: e.target.value })} placeholder="contact@transporteur.com" />
          </Field>
          <FormRow>
            <Field label="N° Teleroute">
              <input className="form-input" value={stForm.numeroTeleroute} onChange={(e) => setStForm({ ...stForm, numeroTeleroute: e.target.value })} placeholder="Numéro Teleroute" />
            </Field>
            <Field label="N° Timocom">
              <input className="form-input" value={stForm.numeroTimocom} onChange={(e) => setStForm({ ...stForm, numeroTimocom: e.target.value })} placeholder="Numéro Timocom" />
            </Field>
          </FormRow>
          <FormRow>
            <Field label="Type de flotte">
              <input className="form-input" value={stForm.flotte} onChange={(e) => setStForm({ ...stForm, flotte: e.target.value })} placeholder="ex : 3 fourgons, 2 porteurs" />
            </Field>
            <Field label="Zone de couverture">
              <input className="form-input" value={stForm.zone} onChange={(e) => setStForm({ ...stForm, zone: e.target.value })} placeholder="ex : Île-de-France, National" />
            </Field>
          </FormRow>
          <div className="flex justify-end gap-2.5 mt-6">
            <button onClick={() => { setStModal(false); setStForm({ id: null, nom: '', siret: '', contact: '', tel: '', email: '', numeroTeleroute: '', numeroTimocom: '', flotte: '', zone: '', typeEntreprise: 'transporteur', numeroTva: '' }); setStSiretCheck(null); }} className="btn-ghost">Annuler</button>
            <button onClick={createSousTraitant} className="px-4 py-2 rounded-md text-[13px] font-semibold bg-[#0d1b2a] text-white hover:bg-[#1a2d3d] transition">{stForm.id ? 'Enregistrer les modifications' : 'Ajouter le sous-traitant'}</button>
          </div>
        </Modal>
      )}

      {/* MODAL CLIENT */}
      {donneurModal && (
        <Modal title="Ajouter un client" onClose={() => { setDonneurModal(false); setTvaCheck(null); }} width={540}>
          <FormRow>
            <Field label="Nom de l'entreprise">
              <input className="form-input" value={donneurForm.nom} onChange={(e) => setDonneurForm({ ...donneurForm, nom: e.target.value })} placeholder="Raison sociale" />
            </Field>
            <Field label="Type">
              <select className="form-input" value={donneurForm.type} onChange={(e) => setDonneurForm({ ...donneurForm, type: e.target.value })}>
                <option>Commissionnaire de transport</option>
                <option>Industriel / Chargeur</option>
                <option>Plateforme logistique</option>
                <option>Autre transporteur</option>
              </select>
            </Field>
          </FormRow>

          <div className="mb-1">
            <label className="text-[11px] uppercase tracking-wide text-[#64748b] mb-1.5 block">Méthode de vérification</label>
            <div className="flex bg-gray-100 rounded-lg p-1 mb-3" style={{ maxWidth: 340 }}>
              <button type="button" onClick={() => setDonneurVerifMethod('tva')} className={`flex-1 py-1.5 rounded-md text-[12px] font-semibold transition ${donneurVerifMethod === 'tva' ? 'bg-[#0d1b2a] text-white' : 'text-[#64748b]'}`}>🇪🇺 Par TVA (recommandé)</button>
              <button type="button" onClick={() => setDonneurVerifMethod('siret')} className={`flex-1 py-1.5 rounded-md text-[12px] font-semibold transition ${donneurVerifMethod === 'siret' ? 'bg-[#0d1b2a] text-white' : 'text-[#64748b]'}`}>🇫🇷 Par SIRET</button>
            </div>
          </div>

          {donneurVerifMethod === 'tva' ? (
            <>
              <Field label="N° TVA intracommunautaire">
                <input className="form-input" value={donneurForm.numeroTva || ''} onChange={(e) => { setDonneurForm({ ...donneurForm, numeroTva: e.target.value }); setTvaCheck(null); }} placeholder="ex : FR12345678901" />
              </Field>
              <div className="mb-4 mt-2 flex items-center gap-2.5 flex-wrap">
                <button
                  type="button"
                  onClick={() => checkTvaLive(donneurForm.numeroTva)}
                  disabled={!donneurForm.numeroTva || tvaCheck?.loading}
                  className="text-[11px] px-3 py-1.5 rounded bg-[#0d1b2a] text-white disabled:opacity-40"
                >
                  {tvaCheck?.loading ? 'Vérification…' : 'Vérifier'}
                </button>
                {tvaCheck && !tvaCheck.loading && tvaCheck.statut === 'valide' && (
                  <Badge cls="bg-emerald-500/15 text-emerald-700">✓ TVA {tvaCheck.pays} valide{tvaCheck.nom ? ` — ${tvaCheck.nom}` : ''}</Badge>
                )}
                {tvaCheck && !tvaCheck.loading && tvaCheck.statut === 'valide' && tvaCheck.adresse && (
                  <span className="text-[10px] text-[#64748b] basis-full">📍 {tvaCheck.adresse}</span>
                )}
                {tvaCheck && !tvaCheck.loading && tvaCheck.statut === 'invalide' && (
                  <Badge cls="bg-red-500/15 text-red-700">✕ Numéro TVA invalide selon le registre européen</Badge>
                )}
                {tvaCheck && !tvaCheck.loading && tvaCheck.statut === 'erreur' && (
                  <Badge cls="bg-amber-400/15 text-amber-700">⚠ Erreur de connexion au registre VIES</Badge>
                )}
              </div>
              <div className="text-[10px] text-[#64748b] mb-4 -mt-2">💡 Le score de solvabilité (basé sur cette vérification TVA) pourra être calculé depuis la liste des clients une fois celui-ci ajouté.</div>
            </>
          ) : (
            <>
              <Field label="SIRET (pour score de solvabilité)">
                <input className="form-input" value={donneurForm.siret || ''} onChange={(e) => setDonneurForm({ ...donneurForm, siret: e.target.value })} placeholder="000 000 000 00000" />
              </Field>
              <div className="mb-4 mt-2 flex items-center gap-2.5 flex-wrap">
                <button
                  type="button"
                  onClick={() => previewSolvabilite(donneurForm.siret)}
                  disabled={!donneurForm.siret}
                  className="text-[11px] px-3 py-1.5 rounded bg-[#0d1b2a] text-white disabled:opacity-40"
                >Vérifier</button>
              </div>
            </>
          )}

          <Field label="Délai de paiement convenu">
            <input className="form-input" value={donneurForm.delai} onChange={(e) => setDonneurForm({ ...donneurForm, delai: e.target.value })} placeholder="ex : 30 jours fin de mois" />
          </Field>
          <div className="flex justify-end gap-2.5 mt-6">
            <button onClick={() => { setDonneurModal(false); setTvaCheck(null); }} className="btn-ghost">Annuler</button>
            <button onClick={createDonneur} className="px-4 py-2 rounded-md text-[13px] font-semibold bg-[#0d1b2a] text-white hover:bg-[#1a2d3d] transition">Ajouter le client</button>
          </div>
        </Modal>
      )}

      {/* MODAL FACTURE */}
      {factureModal && (
        <Modal title="Nouvelle facture" onClose={() => setFactureModal(false)} width={560}>
          <div className="flex bg-gray-100 rounded-lg p-1 mb-5">
            <button
              onClick={() => setFactureForm({ ...factureForm, sens: 'client' })}
              className={`flex-1 py-2.5 rounded-md text-[13px] font-semibold transition ${factureForm.sens === 'client' ? 'bg-sky-500 text-white' : 'text-[#64748b]'}`}
            >📤 Facture client (j'émets)</button>
            <button
              onClick={() => setFactureForm({ ...factureForm, sens: 'fournisseur' })}
              className={`flex-1 py-2.5 rounded-md text-[13px] font-semibold transition ${factureForm.sens === 'fournisseur' ? 'bg-[#0d1b2a] text-white' : 'text-[#64748b]'}`}
            >📥 Facture fournisseur (je reçois)</button>
          </div>

          <FormRow>
            <Field label={factureForm.sens === 'client' ? 'Client (DHL, Sterne, etc.)' : 'Sous-traitant / fournisseur'}>
              <input className="form-input" value={factureForm.tiers} onChange={(e) => setFactureForm({ ...factureForm, tiers: e.target.value })} placeholder="Nom du tiers" />
            </Field>
            <Field label="Mission associée (optionnel)">
              <input className="form-input" value={factureForm.missionRef} onChange={(e) => setFactureForm({ ...factureForm, missionRef: e.target.value })} placeholder="ex : AFFR-2026-219" />
            </Field>
          </FormRow>
          <FormRow>
            <Field label="Date de facture">
              <input type="date" className="form-input" value={factureForm.date} onChange={(e) => setFactureForm({ ...factureForm, date: e.target.value })} />
            </Field>
            <Field label="Date d'échéance">
              <input type="date" className="form-input" value={factureForm.echeance} onChange={(e) => setFactureForm({ ...factureForm, echeance: e.target.value })} />
            </Field>
          </FormRow>
          <Field label="Montant net à payer (€) — sans TVA, transport intracommunautaire">
            <input type="number" className="form-input" value={factureForm.montantHT} onChange={(e) => setFactureForm({ ...factureForm, montantHT: e.target.value })} placeholder="0.00" />
          </Field>
          <Field label="Notes (optionnel)">
            <input className="form-input" value={factureForm.notes} onChange={(e) => setFactureForm({ ...factureForm, notes: e.target.value })} placeholder="Référence, précision..." />
          </Field>

          <div className="bg-white border border-[#d4006e]/30 rounded-lg p-4 mt-4" style={{ background: '#ffffff' }}>
            <div className="flex justify-between font-bold text-[15px]">
              <span>Total à payer</span>
              <span className="text-[#d4006e]" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>
                {fmtEUR(parseFloat(factureForm.montantHT) || 0)}
              </span>
            </div>
            <div className="text-[10px] text-[#64748b] mt-2 italic">{MENTION_TVA}</div>
          </div>

          <div className="flex justify-end gap-2.5 mt-6">
            <button onClick={() => setFactureModal(false)} className="btn-ghost">Annuler</button>
            <button onClick={createFacture} className="px-4 py-2 rounded-md text-[13px] font-semibold bg-[#d4006e] text-white hover:bg-[#b3005d] transition">Créer la facture</button>
          </div>
        </Modal>
      )}

      {/* MODAL ATTESTATION */}
      {attestationModal && (
        <Modal title="Modifier l'attestation" onClose={() => setAttestationModal(false)} width={580}>
          <div className="text-[11px] uppercase tracking-wide text-[#64748b] mb-3">Documents du sous-traitant</div>
          <div className="text-[10px] text-[#64748b] mb-3 -mt-1">💡 Le statut (Manquant / Valide / Expire bientôt / Expiré) est calculé automatiquement à partir des dates ci-dessous — il n'y a rien à sélectionner manuellement. Ces dates sont normalement saisies par le sous-traitant lui-même au dépôt ; corrige-les ici uniquement si nécessaire.</div>

          <FormRow>
            <Field label="Licence — Date de début de validité">
              <input type="date" className="form-input" value={attestationForm.licenceDateDebut} onChange={(e) => setAttestationForm({ ...attestationForm, licenceDateDebut: e.target.value })} />
            </Field>
            <Field label="Licence — Date de fin de validité">
              <input type="date" className="form-input" value={attestationForm.licenceDate} onChange={(e) => setAttestationForm({ ...attestationForm, licenceDate: e.target.value })} />
            </Field>
          </FormRow>

          <FormRow>
            <Field label="Assurance — Date de début de validité">
              <input type="date" className="form-input" value={attestationForm.assuranceDateDebut} onChange={(e) => setAttestationForm({ ...attestationForm, assuranceDateDebut: e.target.value })} />
            </Field>
            <Field label="Assurance — Date de fin de validité">
              <input type="date" className="form-input" value={attestationForm.assuranceDate} onChange={(e) => setAttestationForm({ ...attestationForm, assuranceDate: e.target.value })} />
            </Field>
          </FormRow>

          <FormRow>
            <Field label="K-bis — Date de début de validité">
              <input type="date" className="form-input" value={attestationForm.kbisDateDebut} onChange={(e) => setAttestationForm({ ...attestationForm, kbisDateDebut: e.target.value })} />
            </Field>
            <Field label="K-bis — Date de fin de validité">
              <input type="date" className="form-input" value={attestationForm.kbisDate} onChange={(e) => setAttestationForm({ ...attestationForm, kbisDate: e.target.value })} />
            </Field>
          </FormRow>

          <div className="text-[11px] uppercase tracking-wide text-[#64748b] mt-5 mb-3 pt-4 border-t border-gray-100">Coordonnées du chauffeur</div>
          <FormRow>
            <Field label="Nom du chauffeur">
              <input className="form-input" value={attestationForm.chauffeurNom} onChange={(e) => setAttestationForm({ ...attestationForm, chauffeurNom: e.target.value })} placeholder="Nom et prénom" />
            </Field>
            <Field label="Téléphone">
              <input className="form-input" value={attestationForm.chauffeurTel} onChange={(e) => setAttestationForm({ ...attestationForm, chauffeurTel: e.target.value })} placeholder="06 00 00 00 00" />
            </Field>
          </FormRow>

          <div className="flex justify-end gap-2.5 mt-6">
            <button onClick={() => setAttestationModal(false)} className="btn-ghost">Annuler</button>
            <button onClick={saveAttestation} className="px-4 py-2 rounded-md text-[13px] font-semibold bg-[#d4006e] text-white hover:bg-[#b3005d] transition">Enregistrer</button>
          </div>
        </Modal>
      )}

      <style>{`
        .form-input { width:100%; background:#ffffff; border:1px solid #e2e5ea; border-radius:6px; color:#1c2733; font-size:13px; padding:10px 12px; outline:none; }
        .form-input:focus { border-color:#d4006e; }
        .form-input option { background:#ffffff; color:#1c2733; }
        .btn-ghost { padding:8px 16px; border-radius:6px; font-size:13px; font-weight:600; background:transparent; color:#64748b; border:1px solid #e2e5ea; }
        .btn-ghost:hover { color:#1c2733; border-color:#d4006e; }
        @keyframes fadeIn { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:translateY(0); } }
      `}</style>
    </div>
  );
}

// ===== SUBCOMPONENTS =====

function SectionLabel({ children }) {
  return <div className="text-[9px] uppercase tracking-widest text-white/40 px-5 pt-3 pb-1.5">{children}</div>;
}

function NavItem({ icon, label, active, onClick, badge, badgeColor }) {
  return (
    <div
      onClick={onClick}
      className={`flex items-center gap-2.5 px-5 py-2.5 cursor-pointer text-[13px] font-semibold border-l-[3px] transition ${
        active ? 'text-[#d4006e] bg-[#d4006e]/10 border-[#d4006e]' : 'text-[#3b4fc4] border-transparent hover:text-[#d4006e] hover:bg-[#fdf0f6]'
      }`}
    >
      <span className="text-[16px] w-5 text-center">{icon}</span>
      <span>{label}</span>
      {badge != null && (
        <span className={`ml-auto text-[10px] font-bold px-1.5 py-0.5 rounded-full text-white ${badgeColor || 'bg-[#d4006e]'}`} style={!badgeColor ? {} : { color: '#fff' }}>
          {badge}
        </span>
      )}
    </div>
  );
}

function Card({ title, action, children }) {
  return (
    <div className="bg-white border border-[#d4006e]/30 shadow-[0_4px_18px_rgba(13,27,42,0.12)] rounded-[10px] overflow-hidden" style={{ background: '#ffffff' }}>
      {title && (
        <div className="px-5 py-4 border-b border-[#d4006e]/20 flex items-center justify-between gap-2.5 flex-wrap">
          <div className="text-[14px] font-semibold" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{title}</div>
          {action}
        </div>
      )}
      <div className="p-5">{children}</div>
    </div>
  );
}

function KPI({ label, value, sub, accent, icon }) {
  return (
    <div className="relative bg-white border border-[#d4006e]/30 shadow-[0_4px_18px_rgba(13,27,42,0.12)] rounded-[10px] p-5 overflow-hidden" style={{ background: '#ffffff' }}>
      <div className={`absolute top-0 left-0 right-0 h-[2px] ${accent || 'bg-[#d4006e]'} opacity-70`}></div>
      {icon && <div className="absolute top-4 right-4 text-[22px] opacity-20">{icon}</div>}
      <div className="text-[11px] text-[#64748b] uppercase tracking-wide mb-2">{label}</div>
      <div className="text-[26px] font-bold" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{value}</div>
      {sub && <div className="text-[12px] text-[#64748b] mt-1">{sub}</div>}
    </div>
  );
}

function Badge({ cls, children }) {
  return <span className={`inline-block px-2.5 py-0.5 rounded-full text-[11px] font-semibold whitespace-nowrap ${cls}`}>{children}</span>;
}

function FormRow({ children }) {
  return <div className="grid grid-cols-2 gap-3.5 mb-4">{children}</div>;
}

function Field({ label, children }) {
  return (
    <div>
      <label className="text-[11px] uppercase tracking-wide text-[#64748b] mb-1.5 block">{label}</label>
      {children}
    </div>
  );
}

function Modal({ title, onClose, width, children }) {
  return (
    <div className="fixed inset-0 bg-black/70 z-50 flex items-center justify-center" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="bg-[#ffffff] border border-[#d4006e]/30 shadow-[0_8px_30px_rgba(13,27,42,0.18)] rounded-xl max-h-[88vh] overflow-y-auto" style={{ width: width || 520, maxWidth: '95vw' }}>
        <div className="px-6 py-5 border-b border-[#d4006e]/20 flex justify-between items-center">
          <div className="text-[16px] font-bold" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{title}</div>
          <button onClick={onClose} className="text-[#64748b] text-xl hover:text-white">✕</button>
        </div>
        <div className="p-6">{children}</div>
      </div>
    </div>
  );
}

function FlowBadgeRow({ m }) {
  return m.flux === 'out' ? (
    <Badge cls="bg-[#0d1b2a]/15 text-[#0d1b2a]">▲ Donné</Badge>
  ) : (
    <Badge cls="bg-[#0d1b2a]/15 text-[#0d1b2a]">▼ Reçu</Badge>
  );
}

function DashboardView({ totalOut, totalIn, totalMarge, margeAvgPct, missions, factures, sousTraitants, outCount, inCount, isAdmin }) {
  const today = new Date().toISOString().slice(0, 10);
  const firstOfMonth = new Date(new Date().getFullYear(), new Date().getMonth(), 1).toISOString().slice(0, 10);

  // Chiffres journaliers
  const missionsJour = missions.filter((m) => m.date === today);
  const caJour = missionsJour.filter((m) => m.flux === 'out').reduce((s, m) => s + (m.vendu || 0), 0);
  const margeJour = missionsJour.filter((m) => m.flux === 'out').reduce((s, m) => s + ((m.vendu || 0) - (m.paye || 0)), 0);

  // Chiffres mensuels
  const missionsMois = missions.filter((m) => m.date >= firstOfMonth);
  const caMois = missionsMois.filter((m) => m.flux === 'out').reduce((s, m) => s + (m.vendu || 0), 0);
  const margeMois = missionsMois.filter((m) => m.flux === 'out').reduce((s, m) => s + ((m.vendu || 0) - (m.paye || 0)), 0);
  const facturesEnRetard = factures.filter((f) => f.statut === 'retard').length;
  const aEncaisser = factures.filter((f) => f.sens === 'client' && f.statut !== 'payee').reduce((s, f) => s + (f.montantHT || 0), 0);

  return (
    <div>
      {/* INDICATEUR TEMPS RÉEL */}
      <div className="flex items-center gap-2 mb-4 text-[11px] text-[#64748b]">
        <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse inline-block"></span>
        Tableau de bord connecté en temps réel · Dernière mise à jour : {new Date().toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' })}
      </div>

      {/* CHIFFRES DU JOUR */}
      {isAdmin && (
        <div className="mb-5">
          <div className="text-[11px] uppercase tracking-widest text-[#64748b] mb-2 font-semibold">📅 Aujourd'hui</div>
          <div className="grid grid-cols-4 gap-4">
            <KPI label="Ordres du jour" value={missionsJour.length} sub={missionsJour.length ? `${missionsJour.filter(m=>m.flux==='out').length} transporteur · ${missionsJour.filter(m=>m.flux==='in').length} client` : "Aucun ordre aujourd'hui"} icon="📋" accent="bg-[#d4006e]" />
            <KPI label="CA jour (transporteur)" value={fmtEUR(caJour)} icon="💰" accent="bg-[#d4006e]" />
            <KPI label="Marge du jour" value={fmtEUR(margeJour)} sub={caJour > 0 ? `${((margeJour/caJour)*100).toFixed(1)}% de marge` : ''} icon="📈" accent={margeJour >= 0 ? 'bg-emerald-400' : 'bg-red-400'} />
            <KPI label="Factures en retard" value={facturesEnRetard} sub={facturesEnRetard ? 'À relancer' : 'Aucune'} icon="⚠️" accent={facturesEnRetard ? 'bg-red-400' : 'bg-emerald-400'} />
          </div>
        </div>
      )}

      {/* CHIFFRES DU MOIS */}
      {isAdmin && (
        <div className="mb-5">
          <div className="text-[11px] uppercase tracking-widest text-[#64748b] mb-2 font-semibold">📆 Ce mois ({new Date().toLocaleDateString('fr-FR', { month: 'long', year: 'numeric' })})</div>
          <div className="grid grid-cols-4 gap-4">
            <KPI label="Ordres du mois" value={missionsMois.length} sub={`${missionsMois.filter(m=>m.flux==='out').length} transporteur · ${missionsMois.filter(m=>m.flux==='in').length} client`} icon="📊" accent="bg-[#0d1b2a]" />
            <KPI label="CA mensuel (transporteur)" value={fmtEUR(caMois)} icon="💼" accent="bg-[#0d1b2a]" />
            <KPI label="Marge mensuelle" value={fmtEUR(margeMois)} sub={caMois > 0 ? `${((margeMois/caMois)*100).toFixed(1)}% de marge` : ''} icon="💹" accent={margeMois >= 0 ? 'bg-emerald-400' : 'bg-red-400'} />
            <KPI label="À encaisser (clients)" value={fmtEUR(Math.round(aEncaisser))} sub="Factures non payées" icon="⏳" accent="bg-amber-400" />
          </div>
        </div>
      )}

      {/* GLOBAL */}
      <div className="grid grid-cols-[1fr_auto_1fr] items-center gap-5 bg-white border border-[#d4006e]/20 shadow-[0_4px_18px_rgba(13,27,42,0.10)] rounded-xl px-7 py-5 mb-5">
        <div className="flex items-center gap-4">
          <div className="text-[30px]">▲</div>
          <div>
            <div className="text-[11px] uppercase tracking-wider text-[#64748b] mb-1">Donné (total)</div>
            <div className="text-[24px] font-bold text-[#d4006e]" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{fmtEUR(totalOut)}</div>
            <div className="text-[11px] text-[#64748b] mt-0.5">{outCount} ordres · {sousTraitants.length} transporteurs</div>
          </div>
        </div>
        {isAdmin && (
          <div className="flex flex-col items-center gap-1 px-4">
            <div className="text-[10px] uppercase tracking-wide text-[#64748b]">Marge totale générée</div>
            <div className="text-[15px] font-bold text-[#d4006e]">+ {fmtEUR(totalMarge)}</div>
            <div className="text-[11px] text-[#64748b]">{margeAvgPct.toFixed(1)}% en moyenne</div>
          </div>
        )}
        <div className="flex items-center gap-4 flex-row-reverse text-right">
          <div className="text-[30px]">▼</div>
          <div>
            <div className="text-[11px] uppercase tracking-wider text-[#64748b] mb-1">Reçu (total)</div>
            <div className="text-[24px] font-bold text-[#0d1b2a]" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{fmtEUR(totalIn)}</div>
            <div className="text-[11px] text-[#64748b] mt-0.5">{inCount} ordres reçus</div>
          </div>
        </div>
      </div>

      <Card title="Ordres de transport récents">
        <table className="w-full text-left border-collapse">
          <thead>
            <tr className="text-[10px] uppercase text-[#0d1b2a] bg-[#fce4f0]">
              <th className="py-2.5 pl-3 pb-3 pr-2">N° Ordre</th><th className="py-2.5 pl-3 pb-3 pr-2">Type</th><th className="py-2.5 pl-3 pb-3 pr-2">Trajet</th><th className="py-2.5 pl-3 pb-3 pr-2">Partenaire</th>{isAdmin && <th className="py-2.5 pl-3 pb-3 pr-2">Marge</th>}<th className="py-2.5 pl-3 pb-3">Statut</th>
            </tr>
          </thead>
          <tbody>
            {missions.slice(0, 6).map((m) => {
              const marge = m.flux === 'out' && m.vendu ? ((m.vendu - m.paye) / m.vendu) * 100 : null;
              return (
                <tr key={m.ref} className="border-t border-gray-100">
                  <td className="py-2.5 pr-2 text-[11px] text-[#d4006e] font-semibold">{m.ref}</td>
                  <td className="py-2.5 pr-2"><FlowBadgeRow m={m} /></td>
                  <td className="py-2.5 pr-2 text-[12.5px]">{m.depart} → {m.dest}</td>
                  <td className="py-2.5 pr-2 text-[12.5px]">{m.partenaire}</td>
                  {isAdmin && <td className="py-2.5 pr-2">{marge != null ? <Badge cls={margeClass(marge)}>{marge.toFixed(1)}%</Badge> : <span className="text-[#64748b] text-[11px]">—</span>}</td>}
                  <td className="py-2.5"><Badge cls={STATUS_MAP[m.statut]?.cls}>{STATUS_MAP[m.statut]?.label}</Badge></td>
                </tr>
              );
            })}
            {missions.length === 0 && <tr><td colSpan={6} className="py-6 text-center text-[#64748b] text-[13px]">Aucun ordre de transport — commencez par en créer un !</td></tr>}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ===== GESTION DES RÔLES PAR POSTE =====
// Modules dont l'accès peut être personnalisé individuellement, en plus
// des permissions standards liées au poste.
const MODULES_PERSONNALISABLES = [
  { key: 'marges', label: 'Marges' },
  { key: 'facturation', label: 'Facturation' },
  { key: 'affectation', label: 'Affectation' },
  { key: 'historique-axes', label: 'Axes' },
  { key: 'sous-traitants', label: 'Sous-traitants' },
  { key: 'donneurs', label: 'Clients' },
  { key: 'demandes-devis', label: 'Demandes de devis' },
];

function RolesView({ showToast }) {
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [createModal, setCreateModal] = useState(false);
  const [creating, setCreating] = useState(false);
  const [createForm, setCreateForm] = useState({ nomComplet: '', email: '', password: '', poste: 'exploitant' });
  const [permsModal, setPermsModal] = useState(null); // utilisateur en cours d'édition des droits
  const [permsSelection, setPermsSelection] = useState([]);

  const POSTES = [
    { value: 'admin', label: '👑 Administrateur', desc: 'Accès total — tous les modules' },
    { value: 'exploitant', label: '🚛 Exploitant', desc: 'Ordres de transport, sous-traitants, attestations' },
    { value: 'commercial', label: '💼 Commercial', desc: 'Ordres de transport, affectation, clients' },
    { value: 'comptable', label: '📊 Comptable', desc: 'Facturation uniquement' },
    { value: 'superviseur', label: '👁 Superviseur', desc: 'Lecture seule + marges + axes' },
  ];

  async function loadUsers() {
    const { data, error } = await supabase.from('profiles').select('*').order('created_at');
    if (!error) setUsers(data || []);
    setLoading(false);
  }

  useEffect(() => {
    loadUsers();
  }, []);

  async function updatePoste(userId, newPoste) {
    const newRole = newPoste === 'admin' ? 'admin' : 'exploitant';
    const { error } = await supabase.from('profiles').update({ poste: newPoste, role: newRole }).eq('id', userId);
    if (error) { showToast('⚠️ Erreur mise à jour'); return; }
    setUsers((prev) => prev.map((u) => u.id === userId ? { ...u, poste: newPoste, role: newRole } : u));
    showToast('✅ Poste mis à jour — la modification prend effet à la prochaine connexion');
  }

  async function toggleActif(userId, actifActuel) {
    const { error } = await supabase.from('profiles').update({ actif: !actifActuel }).eq('id', userId);
    if (error) { showToast('⚠️ Erreur'); return; }
    setUsers((prev) => prev.map((u) => u.id === userId ? { ...u, actif: !actifActuel } : u));
    showToast(actifActuel ? '⛔ Compte désactivé' : '✅ Compte réactivé');
  }

  function generatePassword() {
    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
    let pass = '';
    for (let i = 0; i < 10; i++) pass += chars[Math.floor(Math.random() * chars.length)];
    setCreateForm((prev) => ({ ...prev, password: pass }));
  }

  async function handleCreateUser() {
    if (!createForm.email || !createForm.password) {
      showToast('⚠️ Renseigne au moins un e-mail et un mot de passe');
      return;
    }
    setCreating(true);
    const { data, error } = await supabase.functions.invoke('create-user', {
      body: {
        email: createForm.email,
        password: createForm.password,
        nom_complet: createForm.nomComplet || null,
        poste: createForm.poste,
      },
    });
    setCreating(false);
    if (error || data?.error) {
      showToast(`⚠️ ${data?.error || error?.message || 'Erreur lors de la création'}`);
      return;
    }
    await loadUsers();
    showToast(`✅ Compte créé pour ${createForm.email} — communique-lui son mot de passe`);
    setCreateModal(false);
    setCreateForm({ nomComplet: '', email: '', password: '', poste: 'exploitant' });
  }

  function openPermsModal(u) {
    setPermsSelection(Array.isArray(u.permissions_custom) ? u.permissions_custom : []);
    setPermsModal(u);
  }

  function togglePerm(key) {
    setPermsSelection((prev) => prev.includes(key) ? prev.filter((k) => k !== key) : [...prev, key]);
  }

  async function savePermissions() {
    const { error } = await supabase.from('profiles').update({ permissions_custom: permsSelection }).eq('id', permsModal.id);
    if (error) { showToast('⚠️ Erreur lors de la mise à jour des droits'); return; }
    await loadUsers();
    showToast('✅ Droits d\'accès personnalisés enregistrés');
    setPermsModal(null);
  }

  async function resetPermissions() {
    const { error } = await supabase.from('profiles').update({ permissions_custom: null }).eq('id', permsModal.id);
    if (error) { showToast('⚠️ Erreur lors de la réinitialisation'); return; }
    await loadUsers();
    showToast('✅ Droits réinitialisés sur les permissions standards du poste');
    setPermsModal(null);
  }

  return (
    <div>
      <div className="grid grid-cols-5 gap-3 mb-6">
        {POSTES.map((p) => {
          const nb = users.filter((u) => (u.poste || u.role) === p.value).length;
          return (
            <div key={p.value} className="bg-white border border-[#d4006e]/20 rounded-xl p-4 text-center">
              <div className="text-[20px] mb-1">{p.label.split(' ')[0]}</div>
              <div className="text-[12px] font-semibold text-[#0d1b2a]">{p.label.slice(2)}</div>
              <div className="text-[11px] text-[#64748b] mt-1">{p.desc}</div>
              <div className="text-[22px] font-bold text-[#d4006e] mt-2">{nb}</div>
            </div>
          );
        })}
      </div>

      <Card
        title="Membres de l'équipe IBK TMS"
        action={
          <div className="flex items-center gap-2.5">
            <button
              onClick={() => setCreateModal(true)}
              className="px-3.5 py-2 rounded-md text-[12.5px] font-semibold bg-[#d4006e] text-white hover:bg-[#b3005d] transition whitespace-nowrap"
            >
              ➕ Ajouter un utilisateur
            </button>
            <a
              href="https://supabase.com/dashboard/project/fhecosnfbufchvuiufpe/auth/users"
              target="_blank"
              rel="noreferrer"
              className="px-3.5 py-2 rounded-md text-[12.5px] font-semibold bg-[#0d1b2a] text-white hover:bg-[#1a2d3d] transition whitespace-nowrap"
            >
              Gérer / supprimer un compte
            </a>
          </div>
        }
      >
        {loading ? (
          <div className="text-center py-8 text-[#64748b]">Chargement…</div>
        ) : (
          <div className="overflow-x-auto -m-5 p-5">
            <table className="w-full text-left border-collapse min-w-[900px]">
              <thead>
                <tr className="text-[10px] uppercase text-[#0d1b2a] bg-[#fce4f0]">
                  <th className="py-2.5 pl-3 pb-3 pr-3">Utilisateur</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Poste / Rôle</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Accès autorisés</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Statut</th>
                  <th className="py-2.5 pl-3 pb-3">Actions</th>
                </tr>
              </thead>
              <tbody>
                {users.map((u) => {
                  const posteActuel = u.poste || u.role || 'exploitant';
                  const posteInfo = POSTES.find((p) => p.value === posteActuel);
                  const hasCustom = Array.isArray(u.permissions_custom);
                  return (
                    <tr key={u.id} className={`border-t border-gray-100 ${u.actif === false ? 'opacity-50' : ''}`}>
                      <td className="py-3 pr-3">
                        <div className="flex items-center gap-2.5">
                          <div className="rounded-full bg-[#d4006e] text-white flex items-center justify-center font-bold text-[11px]" style={{ width: 30, height: 30 }}>
                            {(u.email || u.id).slice(0, 2).toUpperCase()}
                          </div>
                          <div>
                            <div className="text-[12.5px] font-semibold text-[#0d1b2a]">{u.nom_complet || u.email || u.id.slice(0, 8)}</div>
                            <div className="text-[10px] text-[#64748b]">{u.email || '—'}</div>
                          </div>
                        </div>
                      </td>
                      <td className="py-3 pr-3">
                        <select
                          value={posteActuel}
                          onChange={(e) => updatePoste(u.id, e.target.value)}
                          className="form-input text-[12px] py-1"
                          style={{ width: 190 }}
                        >
                          {POSTES.map((p) => (
                            <option key={p.value} value={p.value}>{p.label}</option>
                          ))}
                        </select>
                      </td>
                      <td className="py-3 pr-3">
                        {hasCustom ? (
                          <div className="flex items-center gap-1.5 flex-wrap max-w-[240px]">
                            <Badge cls="bg-sky-500/15 text-sky-700">Personnalisés</Badge>
                            <span className="text-[10px] text-[#64748b]">{u.permissions_custom.length} module{u.permissions_custom.length > 1 ? 's' : ''}</span>
                          </div>
                        ) : (
                          <div className="text-[11px] text-[#64748b] max-w-[220px]">{posteInfo?.desc || '—'}</div>
                        )}
                      </td>
                      <td className="py-3 pr-3">
                        {u.actif === false
                          ? <Badge cls="bg-red-500/15 text-red-700">⛔ Désactivé</Badge>
                          : <Badge cls="bg-emerald-500/15 text-emerald-700">✓ Actif</Badge>
                        }
                      </td>
                      <td className="py-3">
                        <div className="flex gap-1.5 flex-wrap">
                          <button
                            onClick={() => openPermsModal(u)}
                            className="text-[10px] px-2.5 py-1 rounded border border-[#d4006e]/30 text-[#d4006e] hover:bg-[#fdf0f6] transition"
                          >
                            🔐 Droits d'accès
                          </button>
                          <button
                            onClick={() => toggleActif(u.id, u.actif !== false)}
                            className={`text-[10px] px-2.5 py-1 rounded border transition ${u.actif === false ? 'border-emerald-500/30 text-emerald-700 hover:bg-emerald-50' : 'border-red-500/30 text-red-700 hover:bg-red-50'}`}
                          >
                            {u.actif === false ? '✓ Réactiver' : '⛔ Désactiver'}
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
                {users.length === 0 && (
                  <tr><td colSpan={5} className="py-8 text-center text-[#64748b] text-[13px]">Aucun utilisateur trouvé dans la table profiles.</td></tr>
                )}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      <div className="mt-4 p-4 bg-amber-50 border border-amber-200 rounded-lg text-[12px] text-amber-800">
        ℹ️ <strong>Note :</strong> la modification du poste ou des droits d'accès prend effet à la prochaine connexion de l'utilisateur. La suppression définitive d'un compte se fait toujours depuis le tableau de bord Supabase ("Gérer / supprimer un compte").
      </div>

      {/* MODAL CRÉATION UTILISATEUR */}
      {createModal && (
        <Modal title="Ajouter un utilisateur" onClose={() => setCreateModal(false)} width={480}>
          <Field label="Nom complet">
            <input className="form-input" value={createForm.nomComplet} onChange={(e) => setCreateForm({ ...createForm, nomComplet: e.target.value })} placeholder="Nom et prénom" />
          </Field>
          <div className="h-3"></div>
          <Field label="Adresse e-mail">
            <input type="email" className="form-input" value={createForm.email} onChange={(e) => setCreateForm({ ...createForm, email: e.target.value })} placeholder="prenom@ibkeuroafrique.com" />
          </Field>
          <div className="h-3"></div>
          <Field label="Mot de passe">
            <div className="flex gap-2">
              <input className="form-input" value={createForm.password} onChange={(e) => setCreateForm({ ...createForm, password: e.target.value })} placeholder="Au moins 6 caractères" />
              <button type="button" onClick={generatePassword} className="whitespace-nowrap text-[11px] px-3 py-1.5 rounded border border-[#0d1b2a]/30 text-[#0d1b2a] hover:bg-gray-100">🎲 Générer</button>
            </div>
          </Field>
          <div className="h-3"></div>
          <Field label="Poste">
            <select className="form-input" value={createForm.poste} onChange={(e) => setCreateForm({ ...createForm, poste: e.target.value })}>
              {POSTES.map((p) => (
                <option key={p.value} value={p.value}>{p.label}</option>
              ))}
            </select>
          </Field>
          <div className="text-[10px] text-[#64748b] mt-3">💡 Le compte est créé immédiatement et peut se connecter tout de suite — pense à communiquer l'e-mail et le mot de passe à la personne concernée (WhatsApp, en main propre, etc.), ils ne sont pas envoyés automatiquement.</div>
          <div className="flex justify-end gap-2.5 mt-6">
            <button onClick={() => setCreateModal(false)} className="btn-ghost">Annuler</button>
            <button onClick={handleCreateUser} disabled={creating} className="px-4 py-2 rounded-md text-[13px] font-semibold bg-[#d4006e] text-white hover:bg-[#b3005d] transition disabled:opacity-60">
              {creating ? 'Création…' : 'Créer le compte'}
            </button>
          </div>
        </Modal>
      )}

      {/* MODAL DROITS D'ACCÈS PERSONNALISÉS */}
      {permsModal && (
        <Modal title={`Droits d'accès — ${permsModal.nom_complet || permsModal.email}`} onClose={() => setPermsModal(null)} width={480}>
          <div className="text-[12px] text-[#64748b] mb-4">
            Par défaut, les accès sont déterminés par le poste ({POSTES.find((p) => p.value === (permsModal.poste || permsModal.role))?.label}). Coche les modules ci-dessous pour définir des droits personnalisés qui remplaceront ceux du poste pour cet utilisateur uniquement.
          </div>
          <div className="space-y-2">
            {MODULES_PERSONNALISABLES.map((m) => (
              <label key={m.key} className="flex items-center gap-2.5 p-2.5 rounded-lg border border-gray-200 cursor-pointer hover:bg-gray-50">
                <input type="checkbox" checked={permsSelection.includes(m.key)} onChange={() => togglePerm(m.key)} />
                <span className="text-[13px] text-[#1c2733]">{m.label}</span>
              </label>
            ))}
          </div>
          <div className="text-[10px] text-[#64748b] mt-3">ℹ️ Les modules Tableau de bord, Ordres, Cotation et Ordre de transport restent accessibles à tous les postes et ne sont pas concernés par cette liste. Le module Équipe reste réservé aux administrateurs.</div>
          <div className="flex justify-between gap-2.5 mt-6">
            <button onClick={resetPermissions} className="text-[12px] text-[#64748b] underline hover:text-[#d4006e]">Réinitialiser aux permissions du poste</button>
            <div className="flex gap-2.5">
              <button onClick={() => setPermsModal(null)} className="btn-ghost">Annuler</button>
              <button onClick={savePermissions} className="px-4 py-2 rounded-md text-[13px] font-semibold bg-[#d4006e] text-white hover:bg-[#b3005d] transition">Enregistrer</button>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

// ===== HISTORIQUE DES AXES =====
function HistoriqueAxesView() {
  const [axes, setAxes] = useState([]);
  const [loading, setLoading] = useState(true);
  const [recherche, setRecherche] = useState('');
  const [sortBy, setSortBy] = useState('nb_ordres');

  // Base tarifaire manuelle ("Nouvel axe") : permet aux affréteurs et
  // commerciaux de renseigner un prix de référence pour un trajet, sans
  // attendre qu'une vraie mission ait eu lieu — utile pour négocier.
  const [axesTarifaires, setAxesTarifaires] = useState([]);
  const [loadingTarifs, setLoadingTarifs] = useState(true);
  const [axeModal, setAxeModal] = useState(false);
  const [axeForm, setAxeForm] = useState({ paysDepart: 'France', cpDepart: '', paysArrivee: 'France', cpArrivee: '', prixClient: '', prixTransporteur: '' });
  const [rechercheTarifs, setRechercheTarifs] = useState('');
  const [toast, setToast] = useState('');

  function showToast(msg) {
    setToast(msg);
    setTimeout(() => setToast(''), 3000);
  }

  async function loadAxesTarifaires() {
    const { data, error } = await supabase.from('axes_tarifaires').select('*').order('created_at', { ascending: false });
    if (!error) setAxesTarifaires(data || []);
    setLoadingTarifs(false);
  }

  useEffect(() => {
    loadAxesTarifaires();
  }, []);

  async function createAxe() {
    if (!axeForm.cpDepart || !axeForm.cpArrivee) {
      showToast('⚠️ Renseigne au moins les codes postaux de départ et d\'arrivée');
      return;
    }
    const { error } = await supabase.from('axes_tarifaires').insert({
      pays_depart: axeForm.paysDepart || null,
      cp_depart: axeForm.cpDepart || null,
      pays_arrivee: axeForm.paysArrivee || null,
      cp_arrivee: axeForm.cpArrivee || null,
      prix_client: parseFloat(axeForm.prixClient) || 0,
      prix_transporteur: parseFloat(axeForm.prixTransporteur) || 0,
    });
    if (error) {
      showToast('⚠️ Erreur lors de la création de l\'axe');
      console.error(error);
      return;
    }
    await loadAxesTarifaires();
    setAxeModal(false);
    setAxeForm({ paysDepart: 'France', cpDepart: '', paysArrivee: 'France', cpArrivee: '', prixClient: '', prixTransporteur: '' });
    showToast('✅ Axe tarifaire ajouté');
  }

  async function deleteAxe(id) {
    const { error } = await supabase.from('axes_tarifaires').delete().eq('id', id);
    if (error) {
      showToast('⚠️ Erreur lors de la suppression');
      console.error(error);
      return;
    }
    await loadAxesTarifaires();
    showToast('Axe tarifaire supprimé');
  }

  useEffect(() => {
    async function load() {
      const { data, error } = await supabase
        .from('historique_axes')
        .select('*')
        .order('nb_ordres', { ascending: false });
      if (!error) setAxes(data || []);
      setLoading(false);
    }
    load();
  }, []);

  const filtres = axes.filter((a) => {
    if (!recherche.trim()) return true;
    const q = recherche.toLowerCase();
    return a.depart?.toLowerCase().includes(q) || a.destination?.toLowerCase().includes(q);
  }).sort((a, b) => {
    if (sortBy === 'nb_ordres') return b.nb_ordres - a.nb_ordres;
    if (sortBy === 'prix_moyen') return b.prix_moyen - a.prix_moyen;
    if (sortBy === 'dernier_prix') return b.dernier_prix - a.dernier_prix;
    return 0;
  });

  const totalOrdres = axes.reduce((s, a) => s + (a.nb_ordres || 0), 0);
  const prixMoyenGlobal = axes.length ? axes.reduce((s, a) => s + (a.prix_moyen || 0), 0) / axes.length : 0;
  const axeTop = axes[0];

  return (
    <div>
      <Toast message={toast} />

      <div className="flex items-center justify-between mb-4 flex-wrap gap-3">
        <div>
          <div className="text-[16px] font-bold" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>Base tarifaire (axes de référence)</div>
          <div className="text-[12px] text-[#64748b] mt-0.5">Prix de référence saisis manuellement, pour négocier sans attendre qu'une mission ait eu lieu.</div>
        </div>
        <div className="flex items-center gap-3">
          <div className="relative">
            <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[#64748b] pointer-events-none">🔍</span>
            <input
              className="form-input pl-9"
              style={{ width: 240 }}
              placeholder="Rechercher : pays, code postal..."
              value={rechercheTarifs}
              onChange={(e) => setRechercheTarifs(e.target.value)}
            />
          </div>
          <button onClick={() => setAxeModal(true)} className="px-4 py-2 rounded-md text-[13px] font-semibold bg-[#d4006e] text-white hover:bg-[#b3005d] whitespace-nowrap">+ Nouvel axe</button>
        </div>
      </div>

      <Card>
        {loadingTarifs ? (
          <div className="text-center py-6 text-[#64748b]">Chargement…</div>
        ) : (
          <div className="overflow-x-auto -m-5 p-5">
            <table className="w-full text-left border-collapse min-w-[700px]">
              <thead>
                <tr className="text-[10px] uppercase text-[#0d1b2a] bg-[#fce4f0]">
                  <th className="py-2.5 pl-3 pb-3 pr-3">Départ</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Arrivée</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Prix client</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Prix transporteur</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Marge</th>
                  <th className="py-2.5 pl-3 pb-3"></th>
                </tr>
              </thead>
              <tbody>
                {axesTarifaires.filter((ax) => {
                  if (!rechercheTarifs.trim()) return true;
                  const q = rechercheTarifs.trim().toLowerCase();
                  const champs = [ax.pays_depart, ax.cp_depart, ax.pays_arrivee, ax.cp_arrivee];
                  return champs.some((champ) => champ && champ.toLowerCase().includes(q));
                }).map((ax) => {
                  const marge = (ax.prix_client || 0) - (ax.prix_transporteur || 0);
                  const margePct = ax.prix_client > 0 ? (marge / ax.prix_client) * 100 : 0;
                  return (
                    <tr key={ax.id} className="border-t border-gray-100">
                      <td className="py-2.5 pr-3 font-semibold text-[#0d1b2a]">{ax.cp_depart} {ax.pays_depart}</td>
                      <td className="py-2.5 pr-3 font-semibold text-[#0d1b2a]">{ax.cp_arrivee} {ax.pays_arrivee}</td>
                      <td className="py-2.5 pr-3">{fmtEUR(ax.prix_client)}</td>
                      <td className="py-2.5 pr-3">{fmtEUR(ax.prix_transporteur)}</td>
                      <td className="py-2.5 pr-3"><Badge cls={margeClass(margePct)}>{fmtEUR(marge)} · {margePct.toFixed(1)}%</Badge></td>
                      <td className="py-2.5">
                        <button onClick={() => deleteAxe(ax.id)} className="text-[10px] px-2 py-1 rounded border border-red-500/30 text-red-700 hover:bg-red-500/10">Suppr.</button>
                      </td>
                    </tr>
                  );
                })}
                {axesTarifaires.length === 0 && (
                  <tr><td colSpan={6} className="py-8 text-center text-[#64748b] text-[13px]">Aucun axe tarifaire enregistré — clique sur "+ Nouvel axe" pour en ajouter un.</td></tr>
                )}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      {axeModal && (
        <Modal title="Nouvel axe tarifaire" onClose={() => setAxeModal(false)} width={480}>
          <div className="text-[10px] uppercase tracking-wide text-[#64748b] mb-3 font-semibold">Trajet</div>
          <FormRow>
            <Field label="Pays de départ">
              <input className="form-input" value={axeForm.paysDepart} onChange={(e) => setAxeForm({ ...axeForm, paysDepart: e.target.value })} placeholder="France" />
            </Field>
            <Field label="Code postal de départ">
              <input className="form-input" value={axeForm.cpDepart} onChange={(e) => setAxeForm({ ...axeForm, cpDepart: e.target.value })} placeholder="95140" />
            </Field>
          </FormRow>
          <FormRow>
            <Field label="Pays d'arrivée">
              <input className="form-input" value={axeForm.paysArrivee} onChange={(e) => setAxeForm({ ...axeForm, paysArrivee: e.target.value })} placeholder="Allemagne" />
            </Field>
            <Field label="Code postal d'arrivée">
              <input className="form-input" value={axeForm.cpArrivee} onChange={(e) => setAxeForm({ ...axeForm, cpArrivee: e.target.value })} placeholder="69000" />
            </Field>
          </FormRow>

          <div className="text-[10px] uppercase tracking-wide text-[#64748b] mb-3 mt-3 font-semibold border-t border-gray-100 pt-3">Prix</div>
          <FormRow>
            <Field label="Prix client (€ net)">
              <input type="number" className="form-input" value={axeForm.prixClient} onChange={(e) => setAxeForm({ ...axeForm, prixClient: e.target.value })} placeholder="0.00" />
            </Field>
            <Field label="Prix transporteur (€ net)">
              <input type="number" className="form-input" value={axeForm.prixTransporteur} onChange={(e) => setAxeForm({ ...axeForm, prixTransporteur: e.target.value })} placeholder="0.00" />
            </Field>
          </FormRow>

          <div className="flex justify-end gap-2.5 mt-6">
            <button onClick={() => setAxeModal(false)} className="btn-ghost">Annuler</button>
            <button onClick={createAxe} className="px-4 py-2 rounded-md text-[13px] font-semibold bg-[#d4006e] text-white hover:bg-[#b3005d] transition">Ajouter l'axe</button>
          </div>
        </Modal>
      )}

      <div className="text-[16px] font-bold mt-8 mb-4" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>Historique des axes (généré automatiquement)</div>

      <div className="grid grid-cols-3 gap-4 mb-5">
        <KPI label="Axes enregistrés" value={axes.length} icon="🗺️" accent="bg-[#d4006e]" />
        <KPI label="Total ordres sur ces axes" value={totalOrdres} icon="📋" accent="bg-[#0d1b2a]" />
        <KPI label="Prix moyen global" value={fmtEUR(Math.round(prixMoyenGlobal))} icon="💰" />
      </div>

      {axeTop && (
        <div className="bg-white border border-[#d4006e]/20 rounded-xl p-4 mb-5 flex items-center gap-4">
          <div className="text-[28px]">🏆</div>
          <div>
            <div className="text-[11px] uppercase tracking-wide text-[#64748b] mb-1">Axe le plus fréquent</div>
            <div className="text-[16px] font-bold text-[#0d1b2a]">{axeTop.depart} → {axeTop.destination}</div>
            <div className="text-[12px] text-[#64748b]">{axeTop.nb_ordres} ordres · Prix moyen : {fmtEUR(Math.round(axeTop.prix_moyen))} · Dernier : {fmtEUR(axeTop.dernier_prix)}</div>
          </div>
        </div>
      )}

      <div className="flex items-center justify-between mb-4 flex-wrap gap-3">
        <div className="relative">
          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[#64748b] pointer-events-none">🔍</span>
          <input
            className="form-input pl-9"
            style={{ width: 280 }}
            placeholder="Rechercher par départ ou destination..."
            value={recherche}
            onChange={(e) => setRecherche(e.target.value)}
          />
        </div>
        <div className="flex items-center gap-2">
          <span className="text-[12px] text-[#64748b]">Trier par :</span>
          <select className="form-input" style={{ width: 180 }} value={sortBy} onChange={(e) => setSortBy(e.target.value)}>
            <option value="nb_ordres">Nb d'ordres</option>
            <option value="prix_moyen">Prix moyen</option>
            <option value="dernier_prix">Dernier prix</option>
          </select>
        </div>
      </div>

      <Card title={`Historique des axes (${filtres.length} axes)`}>
        {loading ? (
          <div className="text-center py-8 text-[#64748b]">Chargement…</div>
        ) : (
          <div className="overflow-x-auto -m-5 p-5">
            <table className="w-full text-left border-collapse min-w-[700px]">
              <thead>
                <tr className="text-[10px] uppercase text-[#0d1b2a] bg-[#fce4f0]">
                  <th className="py-2.5 pl-3 pb-3 pr-3">Axe (Départ → Destination)</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Nb ordres</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Prix moyen</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Dernier prix</th>
                  <th className="py-2.5 pl-3 pb-3">Dernière date</th>
                </tr>
              </thead>
              <tbody>
                {filtres.map((a) => (
                  <tr key={a.id} className="border-t border-gray-100">
                    <td className="py-2.5 pr-3 font-semibold text-[#0d1b2a]">
                      <span className="text-[#d4006e]">{a.depart}</span>
                      <span className="text-[#64748b] mx-2">→</span>
                      <span>{a.destination}</span>
                    </td>
                    <td className="py-2.5 pr-3">
                      <Badge cls="bg-[#d4006e]/15 text-[#d4006e]">{a.nb_ordres} ordre{a.nb_ordres > 1 ? 's' : ''}</Badge>
                    </td>
                    <td className="py-2.5 pr-3 font-semibold">{fmtEUR(Math.round(a.prix_moyen))}</td>
                    <td className="py-2.5 pr-3">{fmtEUR(a.dernier_prix)}</td>
                    <td className="py-2.5 text-[11px] text-[#64748b]">{fmtDate(a.derniere_date)}</td>
                  </tr>
                ))}
                {filtres.length === 0 && (
                  <tr><td colSpan={5} className="py-8 text-center text-[#64748b] text-[13px]">
                    {recherche ? `Aucun axe ne correspond à "${recherche}"` : 'Aucun axe enregistré — créez des ordres de transport pour alimenter cet historique automatiquement.'}
                  </td></tr>
                )}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </div>
  );
}

// ===== AFFECTATION — Tableau journalier temps réel =====
function AffectationView({ userEmail, showToast }) {
  const [affectations, setAffectations] = useState([]);
  const [loading, setLoading] = useState(true);
  const [form, setForm] = useState({ numero_ordre: '', affreteur: '', prix_client: '', prix_transporteur: '' });
  const [submitting, setSubmitting] = useState(false);
  const [historique, setHistorique] = useState(false);
  const [dateFiltre, setDateFiltre] = useState('');

  // Chargement des affectations du jour (ou historique)
  const loadAffectations = useCallback(async () => {
    let query = supabase.from('affectations').select('*').order('created_at', { ascending: false });
    if (!historique) {
      // Filtre sur les dernières 24h
      const hier = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();
      query = query.gte('created_at', hier);
    } else if (dateFiltre) {
      // Filtre sur une date précise
      const debut = new Date(dateFiltre + 'T00:00:00').toISOString();
      const fin = new Date(dateFiltre + 'T23:59:59').toISOString();
      query = query.gte('created_at', debut).lte('created_at', fin);
    }
    const { data, error } = await query;
    if (!error) setAffectations(data || []);
    setLoading(false);
  }, [historique, dateFiltre]);

  useEffect(() => {
    loadAffectations();
    // Temps réel : notification dès qu'un collègue ajoute une affectation
    const channel = supabase
      .channel('affectations-realtime')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'affectations' }, (payload) => {
        loadAffectations();
        if (payload.eventType === 'INSERT' && payload.new?.created_by !== userEmail) {
          showToast(`🔔 Nouvelle affectation de ${payload.new.affreteur} — ${payload.new.numero_ordre}`);
        }
      })
      .subscribe();
    return () => { supabase.removeChannel(channel); };
  }, [loadAffectations, userEmail, showToast]);

  async function ajouterAffectation() {
    if (!form.numero_ordre || !form.affreteur) {
      showToast('⚠️ N° ordre et affréteur requis');
      return;
    }
    setSubmitting(true);
    const { error } = await supabase.from('affectations').insert({
      numero_ordre: form.numero_ordre,
      affreteur: form.affreteur,
      prix_client: parseFloat(form.prix_client) || 0,
      prix_transporteur: parseFloat(form.prix_transporteur) || 0,
      created_by: userEmail,
      annule: false,
    });
    setSubmitting(false);
    if (error) { showToast('⚠️ Erreur lors de l\'ajout'); return; }
    setForm({ numero_ordre: '', affreteur: '', prix_client: '', prix_transporteur: '' });
    showToast('✅ Affectation ajoutée');
  }

  async function annulerAffectation(id) {
    const { error } = await supabase.from('affectations').update({ annule: true }).eq('id', id);
    if (error) { showToast('⚠️ Erreur'); return; }
    showToast('Affectation annulée');
  }

  const actives = affectations.filter((a) => !a.annule);
  const totalCA = actives.reduce((s, a) => s + (a.prix_client || 0), 0);
  const totalCout = actives.reduce((s, a) => s + (a.prix_transporteur || 0), 0);
  const totalMarge = totalCA - totalCout;

  return (
    <div>
      {/* Formulaire d'ajout rapide */}
      <Card title="➕ Nouvelle affectation">
        <div className="grid grid-cols-4 gap-3 mb-3">
          <Field label="N° ordre de transport">
            <input className="form-input" value={form.numero_ordre} onChange={(e) => setForm({ ...form, numero_ordre: e.target.value })} placeholder="OT-XXXXXXXX" />
          </Field>
          <Field label="Affréteur / Commercial">
            <input className="form-input" value={form.affreteur} onChange={(e) => setForm({ ...form, affreteur: e.target.value })} placeholder="Nom de l'affréteur" />
          </Field>
          <Field label="Prix client (€ net)">
            <input type="number" className="form-input" value={form.prix_client} onChange={(e) => setForm({ ...form, prix_client: e.target.value })} placeholder="0.00" />
          </Field>
          <Field label="Prix transporteur (€ net)">
            <input type="number" className="form-input" value={form.prix_transporteur} onChange={(e) => setForm({ ...form, prix_transporteur: e.target.value })} placeholder="0.00" />
          </Field>
        </div>
        <div className="flex items-center justify-between">
          {form.prix_client && form.prix_transporteur ? (
            <div className="text-[13px] font-semibold text-[#d4006e]">
              Marge : {fmtEUR((parseFloat(form.prix_client) || 0) - (parseFloat(form.prix_transporteur) || 0))} ({form.prix_client > 0 ? (((parseFloat(form.prix_client) - parseFloat(form.prix_transporteur)) / parseFloat(form.prix_client)) * 100).toFixed(1) : 0}%)
            </div>
          ) : <div />}
          <button onClick={ajouterAffectation} disabled={submitting} className="px-5 py-2 rounded-md text-[13px] font-semibold bg-[#d4006e] text-white hover:bg-[#b3005d] disabled:opacity-60 transition">
            {submitting ? 'Ajout…' : '+ Ajouter l\'affectation'}
          </button>
        </div>
      </Card>

      {/* KPIs du jour */}
      <div className="grid grid-cols-4 gap-4 my-5">
        <KPI label="Affectations du jour" value={actives.length} icon="📋" accent="bg-[#d4006e]" />
        <KPI label="CA total client" value={fmtEUR(totalCA)} icon="💰" />
        <KPI label="Coût transporteurs" value={fmtEUR(totalCout)} icon="🚛" accent="bg-[#0d1b2a]" />
        <KPI label="Marge journalière" value={fmtEUR(totalMarge)} sub={totalCA > 0 ? `${((totalMarge / totalCA) * 100).toFixed(1)}% de marge` : ''} icon="📈" accent={totalMarge >= 0 ? 'bg-emerald-400' : 'bg-red-400'} />
      </div>

      {/* Filtre historique */}
      <div className="flex items-center gap-3 mb-4">
        <button
          onClick={() => { setHistorique(false); setDateFiltre(''); }}
          className={`px-3 py-1.5 rounded-md text-[12px] font-semibold transition ${!historique ? 'bg-[#d4006e] text-white' : 'bg-white border border-[#d4006e]/30 text-[#d4006e]'}`}
        >📅 Aujourd'hui (24h)</button>
        <button
          onClick={() => setHistorique(true)}
          className={`px-3 py-1.5 rounded-md text-[12px] font-semibold transition ${historique ? 'bg-[#0d1b2a] text-white' : 'bg-white border border-[#0d1b2a]/30 text-[#0d1b2a]'}`}
        >🗂 Historique</button>
        {historique && (
          <input type="date" className="form-input" style={{ width: 180 }} value={dateFiltre} onChange={(e) => setDateFiltre(e.target.value)} />
        )}
        <div className="text-[11px] text-[#64748b] ml-auto">
          {!historique ? '🔄 Remise à zéro automatique chaque 24h' : `Affichage historique${dateFiltre ? ` du ${fmtDate(dateFiltre)}` : ''}`}
        </div>
      </div>

      {/* Tableau des affectations */}
      <Card title={`Affectations ${historique ? '— Historique' : '— En cours (24h)'}`}>
        {loading ? (
          <div className="text-center py-8 text-[#64748b]">Chargement…</div>
        ) : (
          <div className="overflow-x-auto -m-5 p-5">
            <table className="w-full text-left border-collapse min-w-[800px]">
              <thead>
                <tr className="text-[10px] uppercase text-[#0d1b2a] bg-[#fce4f0]">
                  <th className="py-2.5 pl-3 pb-3 pr-3">N° Ordre</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Affréteur</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Prix client</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Prix transporteur</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Marge</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Heure</th>
                  <th className="py-2.5 pl-3 pb-3">Statut</th>
                </tr>
              </thead>
              <tbody>
                {affectations.length === 0 && (
                  <tr><td colSpan={7} className="py-8 text-center text-[#64748b] text-[13px]">Aucune affectation pour cette période.</td></tr>
                )}
                {affectations.map((a) => {
                  const marge = (a.prix_client || 0) - (a.prix_transporteur || 0);
                  const margePct = a.prix_client > 0 ? (marge / a.prix_client) * 100 : 0;
                  const heure = new Date(a.created_at).toLocaleTimeString('fr-FR', { hour: '2-digit', minute: '2-digit' });
                  return (
                    <tr key={a.id} className={`border-t border-gray-100 ${a.annule ? 'opacity-40 line-through' : ''}`}>
                      <td className="py-2.5 pr-3 font-semibold text-[#d4006e] text-[12px]" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{a.numero_ordre}</td>
                      <td className="py-2.5 pr-3 text-[12.5px] font-medium">{a.affreteur}</td>
                      <td className="py-2.5 pr-3 font-semibold">{fmtEUR(a.prix_client)}</td>
                      <td className="py-2.5 pr-3">{fmtEUR(a.prix_transporteur)}</td>
                      <td className="py-2.5 pr-3">
                        {a.prix_client > 0 ? (
                          <Badge cls={margeClass(margePct)}>{fmtEUR(marge)} · {margePct.toFixed(1)}%</Badge>
                        ) : <span className="text-[#64748b] text-[11px]">—</span>}
                      </td>
                      <td className="py-2.5 pr-3 text-[11px] text-[#64748b]">{heure}</td>
                      <td className="py-2.5">
                        {a.annule ? (
                          <Badge cls="bg-red-500/15 text-red-700">Annulé</Badge>
                        ) : (
                          <button onClick={() => annulerAffectation(a.id)} className="text-[10px] px-2 py-1 rounded border border-red-500/30 text-red-700 hover:bg-red-500/10">Annuler</button>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              {actives.length > 0 && (
                <tfoot>
                  <tr className="border-t-2 border-[#d4006e]/30 bg-[#fdf0f6]">
                    <td colSpan={2} className="py-2.5 pl-3 text-[11px] font-bold text-[#0d1b2a] uppercase">Total du jour</td>
                    <td className="py-2.5 pr-3 font-bold">{fmtEUR(totalCA)}</td>
                    <td className="py-2.5 pr-3 font-bold">{fmtEUR(totalCout)}</td>
                    <td className="py-2.5 pr-3">
                      <Badge cls={margeClass(totalCA > 0 ? (totalMarge / totalCA) * 100 : 0)}>
                        {fmtEUR(totalMarge)} · {totalCA > 0 ? ((totalMarge / totalCA) * 100).toFixed(1) : 0}%
                      </Badge>
                    </td>
                    <td colSpan={2}></td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
        )}
      </Card>
    </div>
  );
}

function AccesRestreint() {
  return (
    <Card>
      <div className="text-center py-10">
        <div className="text-[36px] mb-3">🔒</div>
        <div className="text-[16px] font-bold text-[#1c2733] mb-1">Accès restreint</div>
        <div className="text-[13px] text-[#64748b]">Cette section est réservée aux administrateurs. Contacte IBK Euro Afrique si tu penses avoir besoin d'y accéder.</div>
      </div>
    </Card>
  );
}

function MissionsView({ missions, flowFilter, setFlowFilter, statusFilter, setStatusFilter, onDelete, onOT, onFacturer, onUpdateStatut, onOpenDocs, isAdmin }) {
  const [recherche, setRecherche] = useState('');

  // Recherche par numéro d'ordre, partenaire (transporteur/client) ou
  // trajet, pour retrouver rapidement un ordre de transport précis.
  const filtres = missions.filter((m) => {
    if (!recherche.trim()) return true;
    const q = recherche.trim().toLowerCase();
    const champs = [m.ref, m.partenaire, m.depart, m.dest];
    return champs.some((champ) => champ && champ.toLowerCase().includes(q));
  });

  return (
    <div>
      <div className="flex justify-between items-center mb-5 flex-wrap gap-2.5">
        <div className="flex gap-2 flex-wrap items-center">
          <div className="relative">
            <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[#64748b] pointer-events-none">🔍</span>
            <input
              className="form-input pl-9"
              style={{ width: 260 }}
              placeholder="Rechercher : n° ordre, partenaire, trajet..."
              value={recherche}
              onChange={(e) => setRecherche(e.target.value)}
            />
          </div>
          <select className="form-input" style={{ width: 190 }} value={flowFilter} onChange={(e) => setFlowFilter(e.target.value)}>
            <option value="all">Transporteur + Client</option>
            <option value="out">▲ Donné</option>
            <option value="in">▼ Reçu</option>
          </select>
          <select className="form-input" style={{ width: 180 }} value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
            <option value="all">Tous les statuts</option>
            <option value="planifiee">Planifié</option>
            <option value="charge">Chargé 📦</option>
            <option value="en_cours">En cours 🚛</option>
            <option value="livree">Délivré ✓</option>
            <option value="annule">Annulé ✕</option>
            <option value="litige">Litige ⚠</option>
          </select>
        </div>
      </div>
      <Card>
        <div className="overflow-x-auto -m-5 p-5">
          <table className="w-full text-left border-collapse min-w-[1000px]">
            <thead>
              <tr className="text-[10px] uppercase text-[#0d1b2a] bg-[#fce4f0]">
                <th className="py-2.5 pl-3 pb-3 pr-3">N° Ordre</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Type</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Partenaire</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Trajet</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Date</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Prix client</th>
                {isAdmin && <th className="py-2.5 pl-3 pb-3 pr-3">Prix transp.</th>}
                {isAdmin && <th className="py-2.5 pl-3 pb-3 pr-3">Marge</th>}
                <th className="py-2.5 pl-3 pb-3 pr-3">Statut</th>
                <th className="py-2.5 pl-3 pb-3"></th>
              </tr>
            </thead>
            <tbody>
              {filtres.map((m) => {
                const marge = m.flux === 'out' && m.vendu ? m.vendu - m.paye : null;
                const margePct = marge != null ? (marge / m.vendu) * 100 : null;
                const isAnnule = m.statut === 'annule';
                return (
                  <tr key={m.ref} className={`border-t border-gray-100 ${isAnnule ? 'opacity-50' : ''}`}>
                    <td className="py-2.5 pr-3 text-[11px] text-[#d4006e] font-semibold">{m.ref}</td>
                    <td className="py-2.5 pr-3"><FlowBadgeRow m={m} /></td>
                    <td className="py-2.5 pr-3 text-[12.5px]">{m.partenaire}</td>
                    <td className="py-2.5 pr-3 text-[12px]">{m.depart} → {m.dest}</td>
                    <td className="py-2.5 pr-3 text-[11px] text-[#64748b]">{fmtDate(m.date)}</td>
                    <td className="py-2.5 pr-3 font-semibold">{fmtEUR(m.vendu)}</td>
                    {isAdmin && <td className="py-2.5 pr-3">{m.flux === 'out' ? fmtEUR(m.paye) : <span className="text-[#64748b]">—</span>}</td>}
                    {isAdmin && <td className="py-2.5 pr-3">{margePct != null ? <Badge cls={margeClass(margePct)}>{fmtEUR(marge)} · {margePct.toFixed(1)}%</Badge> : <span className="text-[#64748b] text-[11px]">N/A</span>}</td>}
                    <td className="py-2.5 pr-3">
                      <select
                        value={m.statut}
                        onChange={(e) => onUpdateStatut(m.ref, e.target.value)}
                        className={`text-[11px] font-semibold rounded-full px-2 py-1 border-0 outline-none cursor-pointer ${STATUS_MAP[m.statut]?.cls}`}
                        style={{ background: 'transparent', minWidth: 110 }}
                      >
                        <option value="planifiee">Planifié</option>
                        <option value="charge">Chargé 📦</option>
                        <option value="en_cours">En cours 🚛</option>
                        <option value="livree">Délivré ✓</option>
                        <option value="annule">Annulé ✕</option>
                        <option value="litige">Litige ⚠</option>
                      </select>
                    </td>
                    <td className="py-2.5 flex gap-1.5 flex-wrap">
                      {m.flux === 'in' && (
                        <button
                          onClick={() => onOpenDocs(m)}
                          className={`text-[10px] px-2 py-1 rounded border ${(m.ordreClientUrl && m.cmrUrl) ? 'border-emerald-500/30 text-emerald-700 hover:bg-emerald-50' : 'border-amber-400/40 text-amber-700 hover:bg-amber-50'}`}
                        >
                          {(m.ordreClientUrl && m.cmrUrl) ? '✓ Documents' : '📎 Documents'}
                        </button>
                      )}
                      {isAdmin && !isAnnule && <button onClick={() => onFacturer(m)} className="text-[10px] px-2 py-1 rounded border border-[#d4006e]/30 text-[#d4006e] hover:bg-[#fdf0f6]">Facturer</button>}
                      <button onClick={() => onOT(m.ref)} className="text-[10px] px-2 py-1 rounded border border-[#0d1b2a]/30 text-[#0d1b2a] hover:bg-gray-100">PDF</button>
                      <button onClick={() => onDelete(m.ref)} className="text-[10px] px-2 py-1 rounded border border-red-500/30 text-red-700 hover:bg-red-500/10">Suppr.</button>
                    </td>
                  </tr>
                );
              })}
              {filtres.length === 0 && (
                <tr><td colSpan={10} className="py-8 text-center text-[#64748b] text-[13px]">
                  {recherche ? `Aucun résultat pour "${recherche}"` : 'Aucun ordre de transport ne correspond à ces filtres.'}
                </td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

function MargesView({ outMissions, totalMarge, margeAvgPct, onDelete }) {
  const sorted = [...outMissions].filter(m => m.vendu).sort((a, b) => ((b.vendu - b.paye) / b.vendu) - ((a.vendu - a.paye) / a.vendu));
  const best = sorted[0];
  const worst = sorted[sorted.length - 1];
  return (
    <div>
      <div className="grid grid-cols-4 gap-4 mb-6">
        <KPI label="Marge totale" value={fmtEUR(totalMarge)} icon="💹" />
        <KPI label="Marge moyenne / mission" value={outMissions.length ? fmtEUR(Math.round(totalMarge / outMissions.length)) : '—'} sub={`soit ${margeAvgPct.toFixed(1)}% du CA affrètement`} />
        {best && <KPI label="Meilleure marge" value={`${(((best.vendu - best.paye) / best.vendu) * 100).toFixed(1)}%`} sub={`${best.ref} · ${best.depart} → ${best.dest}`} accent="bg-emerald-400" />}
        {worst && <KPI label="Marge la plus faible" value={`${(((worst.vendu - worst.paye) / worst.vendu) * 100).toFixed(1)}%`} sub={`${worst.ref} · à surveiller`} accent="bg-red-400" />}
      </div>
      <Card title="Détail marge par mission">
        <div className="overflow-x-auto -m-5 p-5">
          <table className="w-full text-left border-collapse min-w-[800px]">
            <thead>
              <tr className="text-[10px] uppercase text-[#0d1b2a] bg-[#fce4f0]">
                <th className="py-2.5 pl-3 pb-3 pr-3">Référence</th><th className="py-2.5 pl-3 pb-3 pr-3">Trajet</th><th className="py-2.5 pl-3 pb-3 pr-3">Prix vendu</th><th className="py-2.5 pl-3 pb-3 pr-3">Prix payé</th><th className="py-2.5 pl-3 pb-3 pr-3">Marge brute</th><th className="py-2.5 pl-3 pb-3 pr-3">Marge %</th><th className="py-2.5 pl-3 pb-3"></th>
              </tr>
            </thead>
            <tbody>
              {outMissions.map((m) => {
                const marge = m.vendu - m.paye;
                const pct = m.vendu ? (marge / m.vendu) * 100 : 0;
                return (
                  <tr key={m.ref} className="border-t border-gray-100">
                    <td className="py-2.5 pr-3 text-[11px] text-[#d4006e]" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{m.ref}</td>
                    <td className="py-2.5 pr-3 text-[12.5px]">{m.depart} → {m.dest}</td>
                    <td className="py-2.5 pr-3">{fmtEUR(m.vendu)}</td>
                    <td className="py-2.5 pr-3">{fmtEUR(m.paye)}</td>
                    <td className="py-2.5 pr-3 font-bold">{fmtEUR(marge)}</td>
                    <td className="py-2.5 pr-3"><Badge cls={margeClass(pct)}>{pct.toFixed(1)}%</Badge></td>
                    <td className="py-2.5">
                      <button onClick={() => onDelete(m.ref)} title="Supprimer" className="text-[10px] px-2 py-1 rounded border border-red-500/30 text-red-700 hover:bg-red-500/10">🗑️ Suppr.</button>
                    </td>
                  </tr>
                );
              })}
              {outMissions.length === 0 && (
                <tr><td colSpan={7} className="py-8 text-center text-[#64748b] text-[13px]">Aucune mission enregistrée.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

function FacturationView({ factures, onAdd, onUpdateStatut, onDelete, onExportPdf, onUpdateDatePaiement, onValiderComptable }) {
  const [onglet, setOnglet] = React.useState('client');
  const [recherche, setRecherche] = React.useState('');

  const facturesClientAll = factures.filter((f) => f.sens === 'client');
  const facturesFournisseurAll = factures.filter((f) => f.sens === 'fournisseur');

  function matchRecherche(f) {
    if (!recherche.trim()) return true;
    const q = recherche.trim().toLowerCase();
    const champs = [f.numero, f.tiers, f.missionRef];
    return champs.some((champ) => champ && champ.toLowerCase().includes(q));
  }

  const facturesClient = facturesClientAll.filter(matchRecherche);
  const facturesFournisseur = facturesFournisseurAll.filter(matchRecherche);

  const totalEmis = facturesClientAll.reduce((s, f) => s + (f.montantHT || 0), 0);
  const aEncaisser = facturesClientAll.filter((f) => f.statut !== 'payee').reduce((s, f) => s + (f.montantHT || 0), 0);
  const totalDu = facturesFournisseurAll.reduce((s, f) => s + (f.montantHT || 0), 0);
  const enRetard = factures.filter((f) => f.statut === 'retard').length;
  const aValider = facturesFournisseurAll.filter((f) => !f.valideComptable && f.factureUrl).length;

  return (
    <div>
      {/* KPIs */}
      <div className="grid grid-cols-5 gap-4 mb-3">
        <KPI label="Facturé clients" value={fmtEUR(Math.round(totalEmis))} icon="📤" accent="bg-[#d4006e]" />
        <KPI label="À encaisser" value={fmtEUR(Math.round(aEncaisser))} icon="⏳" />
        <KPI label="Dû transporteurs" value={fmtEUR(Math.round(totalDu))} icon="📥" accent="bg-[#0d1b2a]" />
        <KPI label="En retard" value={enRetard} sub={enRetard ? 'À relancer' : 'Aucune'} accent={enRetard ? 'bg-red-400' : 'bg-emerald-400'} />
        <KPI label="À valider comptable" value={aValider} sub="Documents reçus" accent={aValider ? 'bg-amber-400' : 'bg-emerald-400'} />
      </div>
      <div className="text-[11px] text-[#64748b] italic mb-5">{MENTION_TVA}</div>

      {/* Onglets */}
      <div className="flex items-center justify-between mb-5 flex-wrap gap-3">
        <div className="flex bg-white border border-[#d4006e]/20 rounded-lg overflow-hidden">
          <button onClick={() => setOnglet('client')} className={"px-5 py-2.5 text-[13px] font-semibold transition " + (onglet === 'client' ? 'bg-[#d4006e] text-white' : 'text-[#0d1b2a] hover:bg-[#fdf0f6]')}>
            📤 Factures clients ({facturesClientAll.length})
          </button>
          <button onClick={() => setOnglet('fournisseur')} className={"px-5 py-2.5 text-[13px] font-semibold transition " + (onglet === 'fournisseur' ? 'bg-[#0d1b2a] text-white' : 'text-[#0d1b2a] hover:bg-[#fdf0f6]')}>
            📥 Factures transporteurs ({facturesFournisseurAll.length})
          </button>
        </div>
        <div className="flex items-center gap-3">
          <div className="relative">
            <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[#64748b] pointer-events-none">🔍</span>
            <input
              className="form-input pl-9"
              style={{ width: 260 }}
              placeholder="Rechercher : n° facture, tiers, mission..."
              value={recherche}
              onChange={(e) => setRecherche(e.target.value)}
            />
          </div>
          <button onClick={onAdd} className="px-4 py-2 rounded-md text-[13px] font-semibold bg-[#d4006e] text-white hover:bg-[#b3005d] whitespace-nowrap">+ Nouvelle facture</button>
        </div>
      </div>

      {onglet === 'client' && (
        <Card title="Factures clients — Émises">
          <div className="overflow-x-auto -m-5 p-5">
            <table className="w-full text-left border-collapse min-w-[900px]">
              <thead>
                <tr className="text-[10px] uppercase text-[#0d1b2a] bg-[#fce4f0]">
                  <th className="py-2.5 pl-3 pb-3 pr-3">Numéro</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Client</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">N° Ordre</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Date</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Montant net</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Date paiement</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Statut</th>
                  <th className="py-2.5 pl-3 pb-3"></th>
                </tr>
              </thead>
              <tbody>
                {facturesClient.map((f) => {
                  const echeancePassed = f.echeance && new Date(f.echeance) < new Date() && f.statut !== 'payee';
                  return (
                    <tr key={f.id} className={`border-t border-gray-100 ${echeancePassed ? 'bg-red-50' : ''}`}>
                      <td className="py-2.5 pr-3 text-[11px] text-[#d4006e] font-semibold">{f.numero}</td>
                      <td className="py-2.5 pr-3 text-[12.5px] font-medium">{f.tiers}</td>
                      <td className="py-2.5 pr-3 text-[11px] text-[#64748b]">{f.missionRef || '—'}</td>
                      <td className="py-2.5 pr-3 text-[11px] text-[#64748b]">{fmtDate(f.date)}</td>
                      <td className="py-2.5 pr-3 font-semibold">{fmtEUR(f.montantHT)}</td>
                      <td className="py-2.5 pr-3">
                        <input
                          type="date"
                          value={f.datePaiement || ''}
                          onChange={(e) => onUpdateDatePaiement(f.id, e.target.value)}
                          className="form-input text-[11px] py-1 px-2"
                          style={{ width: 130 }}
                          title="Date de paiement réel"
                        />
                      </td>
                      <td className="py-2.5 pr-3">
                        <select value={f.statut} onChange={(e) => onUpdateStatut(f.id, e.target.value)} className={"text-[11px] font-semibold rounded-full px-2.5 py-1 border-0 outline-none cursor-pointer " + (FACTURE_STATUS_MAP[f.statut]?.cls || '')} style={{ background: 'transparent' }}>
                          <option value="emise" className="bg-white text-[#0d1b2a]">Émise</option>
                          <option value="payee" className="bg-white text-[#0d1b2a]">Payée ✓</option>
                          <option value="retard" className="bg-white text-[#0d1b2a]">En retard</option>
                        </select>
                      </td>
                      <td className="py-2.5 flex gap-1.5">
                        <button onClick={() => onExportPdf(f)} className="text-[10px] px-2 py-1 rounded border border-[#d4006e]/30 text-[#d4006e] hover:bg-[#fdf0f6]">PDF</button>
                        <button onClick={() => onDelete(f.id)} className="text-[10px] px-2 py-1 rounded border border-red-500/30 text-red-700 hover:bg-red-500/10">Suppr.</button>
                      </td>
                    </tr>
                  );
                })}
                {facturesClient.length === 0 && <tr><td colSpan={8} className="py-8 text-center text-[#64748b] text-[13px]">{recherche ? `Aucun résultat pour "${recherche}"` : 'Aucune facture client.'}</td></tr>}
              </tbody>
            </table>
          </div>
        </Card>
      )}

      {onglet === 'fournisseur' && (
        <Card title="Factures transporteurs — À valider">
          <div className="text-[12px] text-[#64748b] mb-4 p-3 bg-amber-50 rounded-lg border border-amber-200">
            ⚠️ Le transporteur dépose lui-même sa facture et sa CMR via son lien sécurisé. Le service documentation vérifie, puis la comptable valide et règle la date de paiement.
          </div>
          <div className="overflow-x-auto -m-5 p-5">
            <table className="w-full text-left border-collapse min-w-[1000px]">
              <thead>
                <tr className="text-[10px] uppercase text-[#0d1b2a] bg-[#e8e8f0]">
                  <th className="py-2.5 pl-3 pb-3 pr-3">Numéro</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Transporteur</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">N° Ordre</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Montant net</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Documents</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Délai paiement</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Date paiement</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Validation</th>
                  <th className="py-2.5 pl-3 pb-3 pr-3">Statut</th>
                  <th className="py-2.5 pl-3 pb-3"></th>
                </tr>
              </thead>
              <tbody>
                {facturesFournisseur.map((f) => (
                  <tr key={f.id} className="border-t border-gray-100">
                    <td className="py-2.5 pr-3 text-[11px] text-[#0d1b2a] font-semibold">{f.numero}</td>
                    <td className="py-2.5 pr-3 text-[12.5px] font-medium">{f.tiers}</td>
                    <td className="py-2.5 pr-3 text-[11px] text-[#64748b]">{f.missionRef || '—'}</td>
                    <td className="py-2.5 pr-3 font-semibold">{fmtEUR(f.montantHT)}</td>
                    <td className="py-2.5 pr-3">
                      <div className="flex gap-1.5">
                        {f.factureUrl ? <Badge cls="bg-emerald-500/15 text-emerald-700">✓ Facture</Badge> : <Badge cls="bg-red-500/15 text-red-700">Facture manquante</Badge>}
                        {f.cmrUrl ? <Badge cls="bg-emerald-500/15 text-emerald-700">✓ CMR</Badge> : <Badge cls="bg-red-500/15 text-red-700">CMR manquante</Badge>}
                      </div>
                    </td>
                    <td className="py-2.5 pr-3 text-[12px]">{f.nbJoursPaiement || 30} jours</td>
                    <td className="py-2.5 pr-3">
                      <input type="date" value={f.datePaiement || ''} onChange={(e) => onUpdateDatePaiement(f.id, e.target.value)} className="form-input text-[11px] py-1 px-2" style={{ width: 130 }} />
                    </td>
                    <td className="py-2.5 pr-3">
                      {f.valideComptable ? (
                        <Badge cls="bg-emerald-500/15 text-emerald-700">✓ Validé</Badge>
                      ) : (
                        <button onClick={() => onValiderComptable(f.id)} disabled={!f.factureUrl || !f.cmrUrl} className="text-[10px] px-2 py-1 rounded bg-[#0d1b2a] text-white disabled:opacity-40">Valider</button>
                      )}
                    </td>
                    <td className="py-2.5 pr-3">
                      <select value={f.statut} onChange={(e) => onUpdateStatut(f.id, e.target.value)} className={"text-[11px] font-semibold rounded-full px-2.5 py-1 border-0 outline-none cursor-pointer " + (FACTURE_STATUS_MAP[f.statut]?.cls || '')} style={{ background: 'transparent' }}>
                        <option value="emise" className="bg-white text-[#0d1b2a]">Émise</option>
                        <option value="payee" className="bg-white text-[#0d1b2a]">Payée ✓</option>
                        <option value="retard" className="bg-white text-[#0d1b2a]">En retard</option>
                      </select>
                    </td>
                    <td className="py-2.5">
                      <button onClick={() => onDelete(f.id)} className="text-[10px] px-2 py-1 rounded border border-red-500/30 text-red-700 hover:bg-red-500/10">Suppr.</button>
                    </td>
                  </tr>
                ))}
                {facturesFournisseur.length === 0 && <tr><td colSpan={10} className="py-8 text-center text-[#64748b] text-[13px]">{recherche ? `Aucun résultat pour "${recherche}"` : 'Aucune facture transporteur.'}</td></tr>}
              </tbody>
            </table>
          </div>
        </Card>
      )}
    </div>
  );
}

function SousTraitantsView({ sousTraitants, missions, onAdd, onVerifySiret, onEdit, onDelete, onVerifierTVA }) {
  const [onglet, setOnglet] = React.useState("transporteur");
  const [recherche, setRecherche] = React.useState("");
  const [verifyMenuId, setVerifyMenuId] = React.useState(null);
  const filtres = sousTraitants.filter((s) => {
    if (s.typeEntreprise !== onglet) return false;
    if (!recherche.trim()) return true;
    const q = recherche.toLowerCase();
    return s.nom?.toLowerCase().includes(q) || s.zone?.toLowerCase().includes(q) || s.contact?.toLowerCase().includes(q);
  });
  const nbT = sousTraitants.filter((s) => s.typeEntreprise === "transporteur").length;
  const nbC = sousTraitants.filter((s) => s.typeEntreprise === "commissionnaire").length;
  return (
    <div>
      <div className="flex items-center justify-between mb-5 flex-wrap gap-3">
        <div className="flex items-center gap-3 flex-wrap">
          <div className="flex bg-white border border-[#d4006e]/20 rounded-lg overflow-hidden">
            <button onClick={() => setOnglet("transporteur")} className={"px-4 py-2 text-[13px] font-semibold transition " + (onglet === "transporteur" ? "bg-[#d4006e] text-white" : "text-[#0d1b2a] hover:bg-[#fdf0f6]")}>🚛 Transporteurs <span className="ml-1 text-[11px] opacity-75">({nbT})</span></button>
            <button onClick={() => setOnglet("commissionnaire")} className={"px-4 py-2 text-[13px] font-semibold transition " + (onglet === "commissionnaire" ? "bg-[#0d1b2a] text-white" : "text-[#0d1b2a] hover:bg-[#fdf0f6]")}>🏢 Commissionnaires <span className="ml-1 text-[11px] opacity-75">({nbC})</span></button>
          </div>
          <div className="relative">
            <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[#64748b] pointer-events-none">🔍</span>
            <input className="form-input pl-9" style={{width:240}} placeholder="Rechercher par nom ou axe..." value={recherche} onChange={(e) => setRecherche(e.target.value)} />
          </div>
        </div>
        <button onClick={onAdd} className="px-4 py-2 rounded-md text-[13px] font-semibold bg-[#d4006e] text-white hover:bg-[#b3005d]">+ Ajouter</button>
      </div>
      <div className="grid grid-cols-3 gap-4">
        {filtres.map((s) => {
          const stMissions = missions.filter((m) => m.flux === 'out' && m.partenaire === s.nom);
          const ca = stMissions.reduce((sum, m) => sum + (m.paye || 0), 0);
          const couleur = s.typeEntreprise === 'transporteur' ? '#d4006e' : '#0d1b2a';
          return (
            <div key={s.id} className="bg-white border border-[#d4006e]/20 shadow-[0_4px_18px_rgba(13,27,42,0.10)] rounded-[10px] p-4">
              <div className="flex items-center justify-between gap-2 mb-2.5">
                <div className="flex items-center gap-3">
                  <div className="rounded-lg flex items-center justify-center font-bold text-[13px] flex-shrink-0 text-white" style={{width:38,height:38,background:couleur}}>
                    {s.nom.split(' ').map((n) => n[0]).join('').slice(0, 2)}
                  </div>
                  <div>
                    <div className="font-semibold text-[13px] text-[#0d1b2a]">{s.nom}</div>
                    <div className="text-[11px] text-[#64748b]">{s.contact} · {s.tel}</div>
                  </div>
                </div>
                <div className="flex gap-1 flex-shrink-0">
                  <button onClick={() => onEdit(s)} title="Modifier" className="text-[11px] px-1.5 py-1 rounded border border-[#d4006e]/30 text-[#d4006e] hover:bg-[#fdf0f6]">✏️</button>
                  <button onClick={() => onDelete(s)} title="Supprimer" className="text-[11px] px-1.5 py-1 rounded border border-red-500/30 text-red-700 hover:bg-red-500/10">🗑️</button>
                </div>
              </div>
              <div className="text-[11px] text-[#64748b] mb-1.5">🚛 {s.flotte}</div>
              <div className="text-[11px] text-[#64748b]">📍 {s.zone}</div>
              <div className="mt-2.5 flex items-center gap-2 flex-wrap relative">
                {s.tvaStatut === 'valide' && <Badge cls="bg-emerald-500/15 text-emerald-700">✓ TVA {s.tvaPays} valide</Badge>}
                {s.tvaStatut === 'invalide' && <Badge cls="bg-red-500/15 text-red-700">✕ TVA invalide</Badge>}
                {s.siretStatut === 'actif' && <Badge cls="bg-emerald-500/15 text-emerald-700">✓ SIRET actif</Badge>}
                {s.siretStatut === 'cesse' && <Badge cls="bg-red-500/15 text-red-700">⚠ SIRET radié</Badge>}
                {s.siretStatut === 'introuvable' && <Badge cls="bg-red-500/15 text-red-700">✕ SIRET introuvable</Badge>}
                {!s.tvaStatut && !s.siretStatut && <Badge cls="bg-amber-400/15 text-amber-700">Non vérifié</Badge>}
                <div className="relative">
                  <button
                    onClick={() => setVerifyMenuId(verifyMenuId === s.id ? null : s.id)}
                    className="text-[10px] px-2 py-1 rounded border border-[#d4006e]/30 text-[#d4006e] hover:bg-[#fdf0f6]"
                  >🔍 Vérifier ▾</button>
                  {verifyMenuId === s.id && (
                    <div className="absolute left-0 top-full mt-1 bg-white border border-[#d4006e]/20 rounded-lg shadow-[0_4px_18px_rgba(13,27,42,0.15)] py-1.5 z-20 min-w-[210px]">
                      <button
                        onClick={() => { if (s.numeroTva) onVerifierTVA(s.numeroTva, s.id, 'st'); setVerifyMenuId(null); }}
                        disabled={!s.numeroTva}
                        className="block w-full text-left px-3 py-1.5 text-[12px] text-[#1c2733] hover:bg-[#fdf0f6] disabled:opacity-40 disabled:cursor-not-allowed whitespace-nowrap"
                      >🇪🇺 Vérifier par TVA{!s.numeroTva ? ' (numéro manquant)' : ''}</button>
                      <button
                        onClick={() => { if (s.siret) onVerifySiret(s.siret, s.id); setVerifyMenuId(null); }}
                        disabled={!s.siret}
                        className="block w-full text-left px-3 py-1.5 text-[12px] text-[#1c2733] hover:bg-[#fdf0f6] disabled:opacity-40 disabled:cursor-not-allowed whitespace-nowrap"
                      >🇫🇷 Vérifier par SIRET{!s.siret ? ' (numéro manquant)' : ''}</button>
                    </div>
                  )}
                </div>
              </div>
              {s.siretNomOfficiel && s.siretNomOfficiel !== s.nom && <div className="text-[10px] text-amber-700 mt-1">⚠ Nom officiel : {s.siretNomOfficiel}</div>}
              <div className="grid grid-cols-2 gap-2 mt-2.5 text-[11px]">
                <div><div className="text-[#64748b]">Ordres</div><div className="font-semibold text-[#0d1b2a]">{stMissions.length}</div></div>
                <div><div className="text-[#64748b]">CA généré</div><div className="font-semibold text-[#d4006e]">{fmtEUR(ca)}</div></div>
              </div>
            </div>
          );
        })}
        {filtres.length === 0 && <div className="text-[#64748b] text-[13px] col-span-3 text-center py-10">{recherche ? 'Aucun résultat pour "' + recherche + '"' : 'Aucune entreprise enregistrée dans cet onglet.'}</div>}
      </div>
    </div>
  );
}

function AttestationsView({ attestations, sousTraitants, stNameById, onEdit, onViewDoc, onCopyLink }) {
  const [openMenuId, setOpenMenuId] = useState(null);
  const [openDocsId, setOpenDocsId] = useState(null);
  const [recherche, setRecherche] = useState('');

  function buildLink(a) {
    return `${window.location.origin}${window.location.pathname}?token=${a.publicToken}`;
  }
  function buildMessage(a) {
    return `Bonjour, merci de renseigner les coordonnées de votre chauffeur et de déposer vos documents (licence, assurance, capacité, K-bis, permis) via ce lien IBK Euro Afrique : ${buildLink(a)}`;
  }

  // Recherche par nom d'entreprise, nom du contact/sous-traitant, e-mail
  // (dès qu'il sera disponible côté fiche sous-traitant) et numéro de
  // téléphone, pour retrouver rapidement un dossier de vérification.
  const filtres = attestations.filter((a) => {
    if (!recherche.trim()) return true;
    const q = recherche.trim().toLowerCase();
    const st = sousTraitants?.find((s) => s.id === a.stId);
    const champs = [
      st?.nom,
      st?.contact,
      st?.email,
      st?.tel,
      a.chauffeurNom,
      a.chauffeurTel,
    ];
    return champs.some((champ) => champ && champ.toLowerCase().includes(q));
  });

  return (
    <Card
      title="Suivi documentaire des sous-traitants"
      action={
        <div className="relative">
          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[#64748b] pointer-events-none">🔍</span>
          <input
            className="form-input pl-9"
            style={{ width: 280 }}
            placeholder="Rechercher : entreprise, contact, e-mail, tél..."
            value={recherche}
            onChange={(e) => setRecherche(e.target.value)}
          />
        </div>
      }
    >
      <div className="overflow-x-auto -m-5 p-5">
        <table className="w-full text-left border-collapse min-w-[900px]">
          <thead>
            <tr className="text-[10px] uppercase text-[#0d1b2a] bg-[#fce4f0]">
              <th className="py-2.5 pl-3 pb-3 pr-3">Sous-traitant</th><th className="py-2.5 pl-3 pb-3 pr-3">Chauffeur</th><th className="py-2.5 pl-3 pb-3 pr-3">Licence transport</th><th className="py-2.5 pl-3 pb-3 pr-3">Assurance RC</th><th className="py-2.5 pl-3 pb-3 pr-3">K-bis</th><th className="py-2.5 pl-3 pb-3 pr-3">Statut global</th><th className="py-2.5 pl-3 pb-3"></th>
            </tr>
          </thead>
          <tbody>
            {filtres.map((a) => {
              // Statut de chaque document : calculé à 100% à partir de la
              // présence du fichier et de sa date de fin de validité —
              // aucune intervention humaine, aucune valeur choisie à la main.
              const licenceEff = computeEffectiveStatut(a.licenceDate, !!a.licenceUrl);
              const assuranceEff = computeEffectiveStatut(a.assuranceDate, !!a.assuranceUrl);
              const kbisEff = computeEffectiveStatut(a.kbisDate, !!a.kbisUrl);

              const docs = [
                { label: 'Licence', eff: licenceEff },
                { label: 'Assurance', eff: assuranceEff },
                { label: 'K-bis', eff: kbisEff },
              ];
              const expired = docs.filter((d) => d.eff === 'expire');
              const missing = docs.filter((d) => d.eff === 'manquant');
              const bientot = docs.filter((d) => d.eff === 'expire_bientot');
              const allValid = docs.every((d) => d.eff === 'ok');

              let statutGlobal;
              if (expired.length > 0) {
                statutGlobal = { label: `${expired.map((d) => d.label).join(', ')} expiré${expired.length > 1 ? 's' : ''} — à renouveler`, cls: 'bg-red-500/15 text-red-700' };
              } else if (bientot.length > 0) {
                statutGlobal = { label: `⚠ ${bientot.map((d) => d.label).join(', ')} à renouveler bientôt`, cls: 'bg-amber-400/15 text-amber-700' };
              } else if (missing.length > 0) {
                statutGlobal = { label: missing.length === docs.length ? 'Documents manquants' : `${missing.map((d) => d.label).join(', ')} manquant${missing.length > 1 ? 's' : ''}`, cls: 'bg-gray-200 text-gray-600' };
              } else if (allValid) {
                statutGlobal = { label: '✓ Conforme', cls: 'bg-emerald-500/15 text-emerald-700' };
              } else {
                statutGlobal = { label: '—', cls: 'bg-gray-200 text-gray-600' };
              }

              return (
                <tr key={a.stId} className="border-t border-gray-100">
                  <td className="py-2.5 pr-3 font-semibold">{stNameById(a.stId)}</td>
                  <td className="py-2.5 pr-3">
                    {a.chauffeurNom ? (
                      <>
                        <div className="text-[12.5px] font-medium">{a.chauffeurNom}</div>
                        <div className="text-[10px] text-[#64748b]">{a.chauffeurTel || '—'}</div>
                      </>
                    ) : <span className="text-[#64748b] text-[11px]">Non renseigné</span>}
                  </td>
                  <td className="py-2.5 pr-3">
                    <Badge cls={DOC_BADGE[licenceEff]?.cls}>{DOC_BADGE[licenceEff]?.label}</Badge>
                    {(a.licenceDateDebut || a.licenceDate) && (
                      <div className="text-[10px] text-[#64748b] mt-1">Du : {a.licenceDateDebut ? fmtDate(a.licenceDateDebut) : '—'} au : {a.licenceDate ? fmtDate(a.licenceDate) : '—'}</div>
                    )}
                  </td>
                  <td className="py-2.5 pr-3">
                    <Badge cls={DOC_BADGE[assuranceEff]?.cls}>{DOC_BADGE[assuranceEff]?.label}</Badge>
                    {(a.assuranceDateDebut || a.assuranceDate) && (
                      <div className="text-[10px] text-[#64748b] mt-1">Du : {a.assuranceDateDebut ? fmtDate(a.assuranceDateDebut) : '—'} au : {a.assuranceDate ? fmtDate(a.assuranceDate) : '—'}</div>
                    )}
                  </td>
                  <td className="py-2.5 pr-3">
                    <Badge cls={DOC_BADGE[kbisEff]?.cls}>{DOC_BADGE[kbisEff]?.label}</Badge>
                    {(a.kbisDateDebut || a.kbisDate) && (
                      <div className="text-[10px] text-[#64748b] mt-1">Du : {a.kbisDateDebut ? fmtDate(a.kbisDateDebut) : '—'} au : {a.kbisDate ? fmtDate(a.kbisDate) : '—'}</div>
                    )}
                  </td>
                  <td className="py-2.5 pr-3"><Badge cls={statutGlobal.cls}>{statutGlobal.label}</Badge></td>
                  <td className="py-2.5 relative">
                    <div className="flex gap-1.5">
                      <div className="relative">
                        <button
                          onClick={() => setOpenDocsId(openDocsId === a.stId ? null : a.stId)}
                          className={`text-[10px] px-2 py-1 rounded border whitespace-nowrap ${(a.licenceUrl || a.assuranceUrl || a.kbisUrl) ? 'border-emerald-500/30 text-emerald-700 hover:bg-emerald-50' : 'border-amber-400/40 text-amber-700 hover:bg-amber-50'}`}
                        >📄 Documents ▾</button>
                        {openDocsId === a.stId && (
                          <div className="absolute left-0 top-full mt-1 bg-white border border-[#d4006e]/20 rounded-lg shadow-[0_4px_18px_rgba(13,27,42,0.15)] py-1.5 z-20 min-w-[210px]">
                            {[
                              ['Licence de transport', a.licenceUrl],
                              ['Assurance RC', a.assuranceUrl],
                              ['K-bis', a.kbisUrl],
                            ].map(([label, url]) => (
                              <button
                                key={label}
                                onClick={() => { if (url) { onViewDoc(url); setOpenDocsId(null); } }}
                                disabled={!url}
                                className="flex items-center justify-between w-full text-left px-3 py-1.5 text-[12px] text-[#1c2733] hover:bg-[#fdf0f6] disabled:opacity-40 disabled:cursor-not-allowed whitespace-nowrap gap-3"
                              >
                                <span>{label}</span>
                                <span className={url ? 'text-emerald-600' : 'text-[#94a3b8]'}>{url ? '⬇ Télécharger' : 'Non déposé'}</span>
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                      <button onClick={() => setOpenMenuId(openMenuId === a.stId ? null : a.stId)} className="text-[10px] px-2 py-1 rounded border border-[#0d1b2a]/30 text-[#0d1b2a] hover:bg-[#fdf0f6] whitespace-nowrap">Envoyer le lien ▾</button>
                      <button onClick={() => onEdit(a)} className="text-[10px] px-2 py-1 rounded border border-[#d4006e]/30 text-[#d4006e] hover:bg-[#fdf0f6] whitespace-nowrap">Modifier</button>
                    </div>
                    {openMenuId === a.stId && (
                      <div className="absolute right-0 top-full mt-1 bg-white border border-[#d4006e]/20 rounded-lg shadow-[0_4px_18px_rgba(13,27,42,0.15)] py-1.5 z-20 min-w-[160px]">
                        <a
                          href={`https://wa.me/?text=${encodeURIComponent(buildMessage(a))}`}
                          target="_blank" rel="noreferrer"
                          onClick={() => setOpenMenuId(null)}
                          className="block px-3 py-1.5 text-[12px] text-[#1c2733] hover:bg-[#fdf0f6] whitespace-nowrap"
                        >💬 Par WhatsApp</a>
                        <a
                          href={`mailto:?subject=${encodeURIComponent('Documents sous-traitant — IBK Euro Afrique')}&body=${encodeURIComponent(buildMessage(a))}`}
                          onClick={() => setOpenMenuId(null)}
                          className="block px-3 py-1.5 text-[12px] text-[#1c2733] hover:bg-[#fdf0f6] whitespace-nowrap"
                        >📧 Par email</a>
                        <button
                          onClick={() => { onCopyLink(a); setOpenMenuId(null); }}
                          className="block w-full text-left px-3 py-1.5 text-[12px] text-[#1c2733] hover:bg-[#fdf0f6] whitespace-nowrap"
                        >🔗 Copier le lien</button>
                      </div>
                    )}
                  </td>
                </tr>
              );
            })}
            {filtres.length === 0 && (
              <tr><td colSpan={7} className="py-8 text-center text-[#64748b] text-[13px]">
                {recherche ? `Aucun résultat pour "${recherche}"` : 'Aucun sous-traitant enregistré pour le moment.'}
              </td></tr>
            )}
          </tbody>
        </table>
      </div>
    </Card>
  );
}

function DonneursView({ donneurs, missions, onAdd, onVerifierTVA, onSolvabilite, onSolvabiliteTva, onDelete }) {
  const [recherche, setRecherche] = useState('');

  const filtres = donneurs.filter((d) => {
    if (!recherche.trim()) return true;
    const q = recherche.trim().toLowerCase();
    const champs = [d.nom, d.numeroClient, d.type, d.numeroTva];
    return champs.some((champ) => champ && champ.toLowerCase().includes(q));
  });

  return (
    <div>
      <div className="flex justify-between items-center mb-5 flex-wrap gap-3">
        <div className="text-[13px] text-[#64748b]">Clients pour lesquels IBK agit en tant que transporteur sous-traitant</div>
        <div className="flex items-center gap-3">
          <div className="relative">
            <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[#64748b] pointer-events-none">🔍</span>
            <input
              className="form-input pl-9"
              style={{ width: 260 }}
              placeholder="Rechercher : nom, n° client, TVA..."
              value={recherche}
              onChange={(e) => setRecherche(e.target.value)}
            />
          </div>
          <button onClick={onAdd} className="px-4 py-2 rounded-md text-[13px] font-semibold bg-[#0d1b2a] text-white whitespace-nowrap">+ Ajouter un client</button>
        </div>
      </div>
      <Card>
        <div className="overflow-x-auto -m-5 p-5">
          <table className="w-full text-left border-collapse min-w-[900px]">
            <thead>
              <tr className="text-[10px] uppercase text-[#0d1b2a] bg-[#fce4f0]">
                <th className="py-2.5 pl-3 pb-3 pr-3">N° Client</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Client</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Type</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Ordres reçus</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">CA généré</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">TVA</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Solvabilité</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Délai</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Statut</th>
                <th className="py-2.5 pl-3 pb-3"></th>
              </tr>
            </thead>
            <tbody>
              {filtres.map((d) => {
                const dMissions = missions.filter((m) => m.flux === 'in' && m.partenaire === d.nom);
                const ca = dMissions.reduce((sum, m) => sum + (m.vendu || 0), 0);
                const score = d.scoreSolvabilite;
                const scoreCls = score >= 70 ? 'bg-emerald-500/15 text-emerald-700' : score >= 45 ? 'bg-amber-400/15 text-amber-700' : score != null ? 'bg-red-500/15 text-red-700' : 'bg-gray-100 text-gray-500';
                const scoreMention = score >= 70 ? '🟢 Solide' : score >= 45 ? '🟡 Correct' : score != null ? '🔴 Risqué' : null;
                const scoreSourceTva = d.scoreDetails?.source === 'tva';
                const peutVerifier = !!d.siret || !!d.numeroTva;
                return (
                  <tr key={d.id} className="border-t border-gray-100">
                    <td className="py-2.5 pr-3 text-[11px] font-bold text-[#d4006e]" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{d.numeroClient || '—'}</td>
                    <td className="py-2.5 pr-3 font-semibold text-[#0d1b2a]">{d.nom}</td>
                    <td className="py-2.5 pr-3 text-[11px] text-[#64748b]">{d.type}</td>
                    <td className="py-2.5 pr-3">{dMissions.length}</td>
                    <td className="py-2.5 pr-3 text-[#d4006e] font-semibold">{fmtEUR(ca)}</td>
                    <td className="py-2.5 pr-3">
                      {d.numeroTva ? (
                        <div className="flex flex-col items-start gap-1.5">
                          {d.tvaStatut === 'valide' && <Badge cls="bg-emerald-500/15 text-emerald-700">✓ {d.tvaPays}</Badge>}
                          {d.tvaStatut === 'invalide' && <Badge cls="bg-red-500/15 text-red-700">✕ Invalide</Badge>}
                          {(!d.tvaStatut || d.tvaStatut === 'non_verifie') && <Badge cls="bg-amber-400/15 text-amber-700">Non vérifiée</Badge>}
                          <button onClick={() => onVerifierTVA(d.numeroTva, d.id, 'client')} className="text-[10px] px-2 py-1 rounded border border-[#0d1b2a]/30 text-[#0d1b2a] hover:bg-gray-100 whitespace-nowrap">🇪🇺 Vérifier</button>
                        </div>
                      ) : <span className="text-[11px] text-[#64748b]">—</span>}
                    </td>
                    <td className="py-2.5 pr-3">
                      <div className="flex flex-col items-start gap-1.5">
                        {score != null ? (
                          <Badge cls={scoreCls}>{score}/100 {scoreMention}{scoreSourceTva ? ' (via TVA)' : ''}</Badge>
                        ) : (
                          <span className="text-[11px] text-[#64748b]">—</span>
                        )}
                        <button
                          onClick={() => { if (d.siret) onSolvabilite(d.siret, d.id); else if (d.numeroTva) onSolvabiliteTva(d.numeroTva, d.id); }}
                          disabled={!peutVerifier}
                          className="text-[10px] px-2 py-1 rounded border border-[#d4006e]/30 text-[#d4006e] hover:bg-[#fdf0f6] disabled:opacity-30 disabled:cursor-not-allowed whitespace-nowrap"
                          title={d.siret ? "Calculer le score de solvabilité (via SIRET)" : d.numeroTva ? "Calculer le score de solvabilité (via TVA)" : "Renseigne un SIRET ou un n° de TVA d'abord"}
                        >📊 Vérifier solvabilité</button>
                      </div>
                    </td>
                    <td className="py-2.5 pr-3 text-[12px]">{d.delai || '—'}</td>
                    <td className="py-2.5"><Badge cls={d.statut === 'actif' ? 'bg-emerald-500/15 text-emerald-700' : 'bg-amber-400/15 text-amber-700'}>{d.statut === 'actif' ? 'Actif' : 'Prospect'}</Badge></td>
                    <td className="py-2.5">
                      <button onClick={() => onDelete(d)} title="Supprimer" className="text-[10px] px-2 py-1 rounded border border-red-500/30 text-red-700 hover:bg-red-500/10">🗑️ Suppr.</button>
                    </td>
                  </tr>
                );
              })}
              {filtres.length === 0 && (
                <tr><td colSpan={10} className="py-8 text-center text-[#64748b] text-[13px]">
                  {recherche ? `Aucun résultat pour "${recherche}"` : 'Aucun client enregistré pour le moment.'}
                </td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

// Statuts possibles d'une demande reçue via le site vitrine.
const DEMANDE_STATUT = {
  nouveau: { label: 'Nouveau', cls: 'bg-[#d4006e]/15 text-[#d4006e]' },
  traite: { label: 'Traité', cls: 'bg-emerald-500/15 text-emerald-700' },
  archive: { label: 'Archivé', cls: 'bg-gray-200 text-gray-600' },
};

// Type de demande = quel formulaire du site vitrine a été utilisé.
const DEMANDE_TYPE = {
  contact: { label: '💬 Contact', cls: 'bg-sky-500/15 text-sky-700' },
  transporteur: { label: '🚗 Transporteur', cls: 'bg-emerald-500/15 text-emerald-700' },
  commissionnaire: { label: '🤝 Commissionnaire', cls: 'bg-amber-400/15 text-amber-700' },
  chargeur: { label: '📦 Chargeur', cls: 'bg-[#d4006e]/15 text-[#d4006e]' },
};

// Libellés lisibles pour les champs additionnels stockés dans `details`
// (variables selon le formulaire d'origine : Flotte, Zone, Secteur...).
const DEMANDE_DETAIL_LABELS = {
  Nom: null, Entreprise: null, Contact: null, 'E-mail': null, Téléphone: null,
  Société: null, Profil: null, Message: null, SIRET: 'SIRET',
  Flotte: 'Flotte / véhicules', Zone: 'Zone de couverture',
  Secteur: "Secteur d'activité", Besoin: 'Besoin exprimé',
};

function DemandesDevisView({ demandes, onUpdateStatut, onDelete, onConvert }) {
  const [recherche, setRecherche] = useState('');
  const [expandedId, setExpandedId] = useState(null);

  const filtres = demandes.filter((d) => {
    if (!recherche.trim()) return true;
    const q = recherche.trim().toLowerCase();
    const champs = [d.nom, d.entreprise, d.email, d.telephone];
    return champs.some((champ) => champ && champ.toLowerCase().includes(q));
  });

  const nbNouveau = demandes.filter((d) => d.statut === 'nouveau').length;

  return (
    <div>
      <div className="flex items-center justify-between mb-5 flex-wrap gap-3">
        <div className="text-[13px] text-[#64748b]">
          Demandes reçues depuis le site vitrine (formulaire de contact et IBK Connect)
          {nbNouveau > 0 && <span className="ml-2 font-semibold text-[#d4006e]">— {nbNouveau} nouvelle{nbNouveau > 1 ? 's' : ''}</span>}
        </div>
        <div className="relative">
          <span className="absolute left-3 top-1/2 -translate-y-1/2 text-[#64748b] pointer-events-none">🔍</span>
          <input
            className="form-input pl-9"
            style={{ width: 260 }}
            placeholder="Rechercher : nom, entreprise, e-mail..."
            value={recherche}
            onChange={(e) => setRecherche(e.target.value)}
          />
        </div>
      </div>

      <Card>
        <div className="overflow-x-auto -m-5 p-5">
          <table className="w-full text-left border-collapse min-w-[950px]">
            <thead>
              <tr className="text-[10px] uppercase text-[#0d1b2a] bg-[#fce4f0]">
                <th className="py-2.5 pl-3 pb-3 pr-3">Date</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Type</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Nom / Entreprise</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Coordonnées</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Message</th>
                <th className="py-2.5 pl-3 pb-3 pr-3">Statut</th>
                <th className="py-2.5 pl-3 pb-3"></th>
              </tr>
            </thead>
            <tbody>
              {filtres.map((d) => {
                const extraFields = Object.entries(d.details || {}).filter(
                  ([k, v]) => v && DEMANDE_DETAIL_LABELS[k] !== null && !['Nom', 'Entreprise', 'Contact', 'E-mail', 'Téléphone', 'Société', 'Profil', 'Message'].includes(k)
                );
                return (
                  <React.Fragment key={d.id}>
                    <tr className={`border-t border-gray-100 ${d.statut === 'archive' ? 'opacity-50' : ''}`}>
                      <td className="py-2.5 pr-3 text-[11px] text-[#64748b] whitespace-nowrap">{fmtDate(d.date)}</td>
                      <td className="py-2.5 pr-3"><Badge cls={DEMANDE_TYPE[d.type]?.cls}>{DEMANDE_TYPE[d.type]?.label || d.type}</Badge></td>
                      <td className="py-2.5 pr-3">
                        <div className="text-[12.5px] font-semibold text-[#0d1b2a]">{d.entreprise || d.nom || '—'}</div>
                        {d.entreprise && d.nom && <div className="text-[10px] text-[#64748b]">{d.nom}</div>}
                        {d.profil && <div className="text-[10px] text-[#64748b]">{d.profil}</div>}
                      </td>
                      <td className="py-2.5 pr-3 text-[11.5px]">
                        {d.email && <div><a href={`mailto:${d.email}`} className="text-[#d4006e]">{d.email}</a></div>}
                        {d.telephone && <div className="text-[#64748b]">{d.telephone}</div>}
                      </td>
                      <td className="py-2.5 pr-3 text-[11.5px] text-[#1c2733] max-w-[220px]">
                        <div className="line-clamp-2">{d.message || '—'}</div>
                        {extraFields.length > 0 && (
                          <button onClick={() => setExpandedId(expandedId === d.id ? null : d.id)} className="text-[10px] text-[#d4006e] underline mt-1">
                            {expandedId === d.id ? 'Masquer les détails' : 'Voir les détails'}
                          </button>
                        )}
                      </td>
                      <td className="py-2.5 pr-3">
                        <select
                          value={d.statut}
                          onChange={(e) => onUpdateStatut(d.id, e.target.value)}
                          className={`text-[11px] font-semibold rounded-full px-2.5 py-1 border-0 outline-none cursor-pointer ${DEMANDE_STATUT[d.statut]?.cls}`}
                        >
                          <option value="nouveau">Nouveau</option>
                          <option value="traite">Traité</option>
                          <option value="archive">Archivé</option>
                        </select>
                      </td>
                      <td className="py-2.5">
                        <div className="flex gap-1.5 flex-wrap">
                          {d.entreprise && (
                            <button onClick={() => onConvert(d)} className="text-[10px] px-2 py-1 rounded border border-emerald-500/30 text-emerald-700 hover:bg-emerald-50 whitespace-nowrap">→ Client</button>
                          )}
                          <button onClick={() => onDelete(d)} className="text-[10px] px-2 py-1 rounded border border-red-500/30 text-red-700 hover:bg-red-50">Suppr.</button>
                        </div>
                      </td>
                    </tr>
                    {expandedId === d.id && extraFields.length > 0 && (
                      <tr className="bg-gray-50">
                        <td colSpan={7} className="py-3 px-4">
                          <div className="grid grid-cols-3 gap-3">
                            {extraFields.map(([k, v]) => (
                              <div key={k}>
                                <div className="text-[9px] uppercase text-[#64748b] font-bold">{DEMANDE_DETAIL_LABELS[k] || k}</div>
                                <div className="text-[12px] text-[#1c2733]">{v}</div>
                              </div>
                            ))}
                          </div>
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
              {filtres.length === 0 && (
                <tr><td colSpan={7} className="py-8 text-center text-[#64748b] text-[13px]">
                  {recherche ? `Aucun résultat pour "${recherche}"` : 'Aucune demande reçue pour le moment.'}
                </td></tr>
              )}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

function CotationView({
  cotPaysDepart, setCotPaysDepart, cotCpDepart, setCotCpDepart,
  cotPaysArrivee, setCotPaysArrivee, cotCpArrivee, setCotCpArrivee,
  cotKm, setCotKm, cotTypeCamion, setCotTypeCamion,
  cotPeage, setCotPeage,
  cotBase, cotRevient, cotFinal, margeStandardPct, onUse,
}) {
  return (
    <div className="grid grid-cols-2 gap-5">
      <Card title="🧮 Calcul de prix au kilomètre">
        <div className="text-[10px] uppercase tracking-wide text-[#64748b] mb-3 font-semibold">Trajet (France ou international)</div>
        <FormRow>
          <Field label="Pays de départ">
            <input className="form-input" value={cotPaysDepart} onChange={(e) => setCotPaysDepart(e.target.value)} placeholder="ex : France, Belgique, Espagne..." />
          </Field>
          <Field label="Code postal de départ">
            <input className="form-input" value={cotCpDepart} onChange={(e) => setCotCpDepart(e.target.value)} placeholder="ex : 95140" />
          </Field>
        </FormRow>
        <FormRow>
          <Field label="Pays d'arrivée">
            <input className="form-input" value={cotPaysArrivee} onChange={(e) => setCotPaysArrivee(e.target.value)} placeholder="ex : Allemagne, Italie..." />
          </Field>
          <Field label="Code postal d'arrivée">
            <input className="form-input" value={cotCpArrivee} onChange={(e) => setCotCpArrivee(e.target.value)} placeholder="ex : 69000" />
          </Field>
        </FormRow>

        <div className="text-[10px] uppercase tracking-wide text-[#64748b] mb-3 mt-2 font-semibold border-t border-gray-100 pt-3">Véhicule & distance</div>
        <FormRow>
          <Field label="Distance (km)"><input type="number" className="form-input" value={cotKm} onChange={(e) => setCotKm(e.target.value)} placeholder="450" /></Field>
          <Field label="Type de camion">
            <select className="form-input" value={cotTypeCamion} onChange={(e) => setCotTypeCamion(e.target.value)}>
              <option>Tautliner</option>
              <option>Mega Tautliner</option>
              <option>Box</option>
              <option>Frigo</option>
            </select>
          </Field>
        </FormRow>
        <Field label="Frais de péage estimés (€)"><input type="number" className="form-input" value={cotPeage} onChange={(e) => setCotPeage(e.target.value)} /></Field>

        <div className="bg-white border border-[#d4006e]/30 rounded-lg p-4.5 mt-4" style={{ background: '#ffffff' }}>
          <Row label="Coût de transport de base" val={fmtEUR(cotBase)} />
          <Row label="Péages" val={fmtEUR(parseFloat(cotPeage) || 0)} />
          <Row label="Coût de revient total" val={fmtEUR(cotRevient)} />
          <div className="flex justify-between pt-2.5 mt-1.5 border-t border-[#d4006e]/20 font-bold text-[15px]">
            <span>Prix de vente conseillé (net, sans TVA)</span>
            <span className="text-[#d4006e]" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{fmtEUR(Math.round(cotFinal))}</span>
          </div>
          <div className="text-[10px] text-[#64748b] mt-1.5">Marge standard de {margeStandardPct}% déjà incluse dans ce prix.</div>
        </div>
        <button onClick={onUse} className="w-full mt-3.5 px-4 py-2.5 rounded-md text-[13px] font-semibold bg-[#d4006e] text-white">Utiliser ce prix dans une nouvelle mission</button>
      </Card>

      <Card title="Grilles tarifaires de référence">
        <table className="w-full text-left border-collapse">
          <thead><tr className="text-[10px] uppercase text-[#0d1b2a] bg-[#fce4f0]"><th className="py-2.5 pl-3 pb-3 pr-3">Type de camion</th><th className="py-2.5 pl-3 pb-3 pr-3">€/km</th><th className="py-2.5 pl-3 pb-3">Usage type</th></tr></thead>
          <tbody>
            <Tr v="Box" p="1,10 €" u="Messagerie / colis palettisés" />
            <Tr v="Tautliner" p="1,15 €" u="Marchandise bâchée standard" />
            <Tr v="Frigo" p="1,30 €" u="Température dirigée" />
            <Tr v="Mega Tautliner" p="1,45 €" u="Grand volume / longue distance" />
          </tbody>
        </table>
        <div className="mt-4 p-3.5 rounded-lg text-[11px] text-[#64748b]" style={{ background: 'rgba(201,168,76,0.08)' }}>
          💡 Grille indicative — à ajuster selon la conjoncture carburant et la disponibilité des bourses de fret. Ces tarifs par défaut peuvent être modifiés directement dans le code si besoin d'un ajustement précis.
        </div>
      </Card>
    </div>
  );
}

function Row({ label, val }) {
  return <div className="flex justify-between py-1.5 text-[13px]"><span>{label}</span><span className="font-semibold" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>{val}</span></div>;
}
function Tr({ v, p, u }) {
  return (
    <tr className="border-t border-gray-100">
      <td className="py-2.5 pr-3">{v}</td>
      <td className="py-2.5 pr-3 font-semibold text-[#d4006e]">{p}</td>
      <td className="py-2.5 text-[11px] text-[#64748b]">{u}</td>
    </tr>
  );
}

// Champ Transporteur avec recherche intelligente : dès 2 lettres saisies,
// propose automatiquement tous les sous-traitants dont le nom correspond.
// En dessous de 2 caractères (ou champ vide au focus), affiche la liste
// complète — combine donc liste déroulante classique et auto-complétion,
// conformément au cahier des charges.
function TransporteurAutocomplete({ value, onChange, sousTraitants }) {
  const [open, setOpen] = useState(false);

  const query = (value || '').trim().toLowerCase();
  const showFiltered = query.length >= 2;
  const list = showFiltered
    ? (sousTraitants || []).filter((s) => s.nom.toLowerCase().includes(query))
    : (sousTraitants || []);

  function selectOption(nom) {
    onChange(nom);
    setOpen(false);
  }

  return (
    <div className="relative">
      <input
        className="form-input"
        value={value || ''}
        onChange={(e) => { onChange(e.target.value); setOpen(true); }}
        onFocus={() => setOpen(true)}
        onBlur={() => setTimeout(() => setOpen(false), 150)}
        placeholder="Rechercher ou sélectionner un transporteur..."
        autoComplete="off"
      />
      {open && list.length > 0 && (
        <div className="absolute left-0 top-full mt-1 bg-white border border-[#d4006e]/20 rounded-lg shadow-[0_4px_18px_rgba(13,27,42,0.15)] py-1.5 z-30 max-h-56 overflow-y-auto w-full">
          {list.map((s) => (
            <button
              key={s.id}
              type="button"
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => selectOption(s.nom)}
              className="block w-full text-left px-3 py-1.5 text-[13px] text-[#1c2733] hover:bg-[#fdf0f6]"
            >
              {s.nom}
            </button>
          ))}
        </div>
      )}
      {open && showFiltered && list.length === 0 && (
        <div className="absolute left-0 top-full mt-1 bg-white border border-[#d4006e]/20 rounded-lg shadow-[0_4px_18px_rgba(13,27,42,0.15)] py-2 px-3 z-30 w-full text-[12px] text-[#64748b]">
          Aucun sous-traitant trouvé pour "{value}"
        </div>
      )}
    </div>
  );
}

function OrdreTransportView({ lv, setLv, sousTraitants, otExistingRef, onGenerate, onReset }) {
  const upd = (k) => (e) => setLv({ ...lv, [k]: e.target.value });
  return (
    <div className="grid grid-cols-2 gap-5">
      <Card title="📄 Transport Order">
        <div className="flex items-center justify-between mb-4 p-3 rounded-lg bg-gray-50 border border-gray-200">
          <div>
            <div className="text-[10px] uppercase tracking-wide text-[#64748b]">Mission Number</div>
            <div className="text-[15px] font-bold text-[#0d1b2a]" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>
              {lv.ref || 'Will be generated automatically upon creation'}
            </div>
          </div>
          {otExistingRef && (
            <button onClick={onReset} className="text-[11px] px-3 py-1.5 rounded border border-[#d4006e]/30 text-[#d4006e] hover:bg-[#fdf0f6]">+ New Mission</button>
          )}
        </div>

        <FormRow>
          <Field label="Pickup Date"><input type="date" className="form-input" value={lv.date} onChange={upd('date')} /></Field>
          <Field label="Delivery Date"><input type="date" className="form-input" value={lv.dateLivraison} onChange={upd('dateLivraison')} /></Field>
        </FormRow>

        <div className="text-[10px] uppercase tracking-wide text-[#64748b] mb-3 mt-2 font-semibold border-t border-gray-100 pt-3">IBK Freight Broker (Negotiation)</div>
        <div className="grid grid-cols-3 gap-3.5 mb-4">
          <Field label="Name"><input className="form-input" value={lv.affreteur} onChange={upd('affreteur')} placeholder="IBK broker's name" /></Field>
          <Field label="Email"><input type="email" className="form-input" value={lv.affreteurEmail} onChange={upd('affreteurEmail')} placeholder="email@ibkeuroafrique.com" /></Field>
          <Field label="Phone"><input className="form-input" value={lv.affreteurTel} onChange={upd('affreteurTel')} placeholder="06 00 00 00 00" /></Field>
        </div>

        <div className="text-[10px] uppercase tracking-wide text-[#64748b] mb-3 mt-2 font-semibold border-t border-gray-100 pt-3">Carrier Performing the Transport</div>
        <FormRow>
          <Field label="Carrier (registered subcontractor)">
            <TransporteurAutocomplete
              value={lv.transporteur}
              onChange={(nom) => setLv({ ...lv, transporteur: nom })}
              sousTraitants={sousTraitants}
            />
          </Field>
          <Field label="Vehicle Registration"><input className="form-input" value={lv.plaque} onChange={upd('plaque')} placeholder="AB-123-CD" /></Field>
        </FormRow>
        <Field label="Company Address">
          <input className="form-input" value={lv.transporteurAdresse} onChange={upd('transporteurAdresse')} placeholder="Carrier's full address" />
        </Field>
        <div className="grid grid-cols-3 gap-3.5 mb-4 mt-4">
          <Field label="Contact at Carrier"><input className="form-input" value={lv.transporteurContact} onChange={upd('transporteurContact')} placeholder="Name of the contact person" /></Field>
          <Field label="Contact Email"><input type="email" className="form-input" value={lv.transporteurContactEmail} onChange={upd('transporteurContactEmail')} placeholder="email@carrier.com" /></Field>
          <Field label="Contact Phone"><input className="form-input" value={lv.transporteurContactTel} onChange={upd('transporteurContactTel')} placeholder="06 00 00 00 00" /></Field>
        </div>

        <div className="text-[10px] uppercase tracking-wide text-[#64748b] mb-3 mt-2 font-semibold border-t border-gray-100 pt-3">Loading</div>
        <div className="grid grid-cols-3 gap-3.5 mb-4">
          <Field label="Loading Location"><input className="form-input" value={lv.lieuCharge} onChange={upd('lieuCharge')} placeholder="Loading address" /></Field>
          <Field label="Postal Code"><input className="form-input" value={lv.cpCharge} onChange={upd('cpCharge')} placeholder="95140" /></Field>
          <Field label="Country"><input className="form-input" value={lv.paysCharge} onChange={upd('paysCharge')} placeholder="France" /></Field>
        </div>
        <div className="grid grid-cols-3 gap-3.5 mb-4">
          <Field label="Loading Client"><input className="form-input" value={lv.clientChargement} onChange={upd('clientChargement')} placeholder="Site/contact name" /></Field>
          <Field label="Start Time"><input type="time" className="form-input" value={lv.heureChargementDebut} onChange={upd('heureChargementDebut')} /></Field>
          <Field label="End Time"><input type="time" className="form-input" value={lv.heureChargementFin} onChange={upd('heureChargementFin')} /></Field>
        </div>

        <div className="text-[10px] uppercase tracking-wide text-[#64748b] mb-3 mt-2 font-semibold border-t border-gray-100 pt-3">Delivery</div>
        <div className="grid grid-cols-3 gap-3.5 mb-4">
          <Field label="Delivery Location"><input className="form-input" value={lv.lieuLivre} onChange={upd('lieuLivre')} placeholder="Delivery address" /></Field>
          <Field label="Postal Code"><input className="form-input" value={lv.cpLivre} onChange={upd('cpLivre')} placeholder="69000" /></Field>
          <Field label="Country"><input className="form-input" value={lv.paysLivre} onChange={upd('paysLivre')} placeholder="France" /></Field>
        </div>
        <div className="grid grid-cols-3 gap-3.5 mb-4">
          <Field label="Delivery Client"><input className="form-input" value={lv.clientLivraison} onChange={upd('clientLivraison')} placeholder="Site/contact name" /></Field>
          <Field label="Start Time"><input type="time" className="form-input" value={lv.heureLivraisonDebut} onChange={upd('heureLivraisonDebut')} /></Field>
          <Field label="End Time"><input type="time" className="form-input" value={lv.heureLivraisonFin} onChange={upd('heureLivraisonFin')} /></Field>
        </div>

        <div className="text-[10px] uppercase tracking-wide text-[#64748b] mb-3 mt-2 font-semibold border-t border-gray-100 pt-3">Goods</div>
        <div className="grid grid-cols-3 gap-3.5 mb-4">
          <Field label="Nature of Goods"><input className="form-input" value={lv.marchandise} onChange={upd('marchandise')} placeholder="e.g. Parcels, pallets..." /></Field>
          <Field label="Weight"><input className="form-input" value={lv.poids} onChange={upd('poids')} placeholder="e.g. 850 kg" /></Field>
          <Field label="LDM (Linear Meters)"><input className="form-input" value={lv.ldm} onChange={upd('ldm')} placeholder="e.g. 2.4 lm" /></Field>
        </div>
        <Field label="Pallet Exchange">
          <select className="form-input" value={lv.exchangePalettes ? 'oui' : 'non'} onChange={(e) => setLv({ ...lv, exchangePalettes: e.target.value === 'oui' })}>
            <option value="non">No</option>
            <option value="oui">Yes</option>
          </select>
        </Field>

        <div className="text-[10px] uppercase tracking-wide text-[#64748b] mb-3 mt-4 font-semibold border-t border-gray-100 pt-3">Price & Payment</div>

        <div className="mb-4 p-3 rounded-lg bg-amber-50 border border-amber-200">
          <Field label="🔒 Client Price (€ net) — internal use only">
            <input type="number" className="form-input" value={lv.prixClient} onChange={upd('prixClient')} placeholder="0.00" />
          </Field>
          <div className="text-[10px] text-amber-800 mt-1.5">This price feeds Margins, Invoicing and the Dashboard — it never appears on the PDF given to the carrier.</div>
        </div>

        <FormRow>
          <Field label="Agreed Carrier Price (€ net, VAT excl.)"><input type="number" className="form-input" value={lv.prix} onChange={upd('prix')} placeholder="0.00" /></Field>
          <Field label="Agreed Payment Terms">
            <select className="form-input" value={lv.delaiPaiement} onChange={upd('delaiPaiement')}>
              <option>30 days end of month</option><option>45 days end of month</option><option>60 days end of month</option><option>Cash</option><option>15 days</option>
            </select>
          </Field>
        </FormRow>

        {lv.prixClient && lv.prix && (
          <div className="bg-white border border-[#d4006e]/30 rounded-lg p-3.5 mb-4">
            <div className="flex justify-between font-bold text-[14px]">
              <span>Estimated Gross Margin</span>
              <span className="text-[#d4006e]">
                {(() => { const v = parseFloat(lv.prixClient) || 0; const p = parseFloat(lv.prix) || 0; const pct = v > 0 ? ((v - p) / v) * 100 : 0; return `${(v - p).toFixed(2)} € (${pct.toFixed(1)}%)`; })()}
              </span>
            </div>
          </div>
        )}

        <button onClick={onGenerate} className="w-full mt-3.5 px-4 py-2.5 rounded-md text-[13px] font-semibold bg-[#d4006e] text-white">
          {otExistingRef ? '⬇ Regenerate PDF' : '✓ Create Mission & Download PDF'}
        </button>
      </Card>

      <Card title="Preview">
        <div className="bg-white text-[#1a1a1a] rounded-lg p-7 text-[11px] leading-relaxed">
          <div className="flex justify-between border-b-[3px] border-[#0d1b2a] pb-3.5 mb-4">
            <div>
              <div className="font-bold text-[16px] text-[#0d1b2a]" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>IBK EURO AFRIQUE</div>
              <div className="text-[10px] text-gray-600">17-19 Boulevard de la Muette<br />95140 Garges-lès-Gonesse<br />SIRET 902 519 743 00014</div>
            </div>
            <div className="text-right">
              <div className="text-[18px] text-[#0d1b2a] font-bold">TRANSPORT ORDER</div>
              <div className="text-[10px] text-gray-600">No. {lv.ref || '—'}</div>
              <div className="text-[10px] text-gray-600">Pickup: {lv.date ? fmtDate(lv.date) : '—'}</div>
              <div className="text-[10px] text-gray-600">Delivery: {lv.dateLivraison ? fmtDate(lv.dateLivraison) : '—'}</div>
            </div>
          </div>
          <div className="border border-gray-300 rounded-md p-2.5 mb-4">
            <div className="text-[9px] uppercase text-gray-500 font-bold mb-1.5">IBK Freight Broker</div>
            <div>{lv.affreteur || '—'}</div>
            <div className="text-[10px] text-gray-500">{lv.affreteurEmail || '—'} · {lv.affreteurTel || '—'}</div>
          </div>
          <div className="border border-gray-300 rounded-md p-2.5 mb-4">
            <div className="text-[9px] uppercase text-gray-500 font-bold mb-1.5">Carrier</div>
            <div>{lv.transporteur || '—'}</div>
            <div className="text-[10px] text-gray-500">{lv.transporteurAdresse || '—'}</div>
            <div className="text-[10px] text-gray-500 mt-1">Contact: {lv.transporteurContact || '—'} · {lv.transporteurContactEmail || '—'} · {lv.transporteurContactTel || '—'}</div>
            <div className="mt-1"><b>Vehicle:</b> {lv.plaque || '—'}</div>
          </div>
          <div className="grid grid-cols-2 gap-4 mb-4">
            <PreviewBox title="Loading" content={`${lv.lieuCharge || '—'} (${lv.cpCharge || '—'} ${lv.paysCharge || ''})`} />
            <PreviewBox title="Delivery" content={`${lv.lieuLivre || '—'} (${lv.cpLivre || '—'} ${lv.paysLivre || ''})`} />
          </div>
          <div className="border border-gray-300 rounded-md p-2.5 mb-4">
            <div className="text-[9px] uppercase text-gray-500 font-bold mb-1.5">Goods Transported</div>
            <div><b>Nature:</b> {lv.marchandise || '—'}</div>
            <div><b>Weight:</b> {lv.poids || '—'}</div>
            <div><b>LDM:</b> {lv.ldm || '—'}</div>
            <div><b>Pallet Exchange:</b> {lv.exchangePalettes ? 'Yes' : 'No'}</div>
          </div>
          <div className="border border-gray-300 rounded-md p-2.5">
            <div className="text-[9px] uppercase text-gray-500 font-bold mb-1.5">Agreed Price</div>
            <div className="text-[16px] font-bold text-[#0d1b2a]">{(parseFloat(lv.prix) || 0).toFixed(2)} € net</div>
          </div>
        </div>
      </Card>
    </div>
  );
}

function PreviewBox({ title, content }) {
  return (
    <div className="border border-gray-300 rounded-md p-2.5">
      <div className="text-[9px] uppercase text-gray-500 font-bold mb-1.5">{title}</div>
      <div>{content || '—'}</div>
    </div>
  );
}

// Champ de dépôt de fichier avec confirmation visuelle explicite du nom du
// fichier sélectionné. Nécessaire car le style générique .form-input
// (padding, fond) écrase le rendu natif du sélecteur de fichier du
// navigateur, qui masque parfois le nom du fichier une fois choisi —
// l'utilisateur ne voyait donc aucune confirmation que son fichier avait
// bien été pris en compte avant l'envoi.
function FileField({ label, file, onChange, accept = '.pdf,image/*' }) {
  return (
    <Field label={label}>
      <input type="file" accept={accept} className="form-input" onChange={onChange} />
      {file ? (
        <div className="mt-1.5 flex items-center gap-1.5 text-[11px] text-emerald-700 font-medium">
          <span>✓</span><span className="truncate">{file.name}</span>
        </div>
      ) : (
        <div className="mt-1.5 text-[11px] text-[#94a3b8]">Aucun fichier sélectionné</div>
      )}
    </Field>
  );
}

// ===== POINT D'ENTRÉE : gère la session de connexion =====
// Tant que personne n'est connecté, on affiche l'écran de Login.
// Une fois connecté, on affiche l'application (AppShell) avec
// toutes les fonctionnalités.
// Exception : si l'URL contient ?token=..., on affiche le formulaire
// public destiné au sous-traitant, sans exiger de connexion.
export default function App() {
  const [session, setSession] = useState(undefined);
  const [role, setRole] = useState(null);
  const [poste, setPoste] = useState(null);
  const [profil, setProfil] = useState(null);

  const publicToken = new URLSearchParams(window.location.search).get('token');

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, newSession) => {
      setSession(newSession);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    async function loadRole() {
      if (!session) { setRole(null); setPoste(null); setProfil(null); return; }
      const { data } = await supabase.from('profiles').select('role, poste, nom_complet, telephone, actif, permissions_custom').eq('id', session.user.id).single();
      setRole(data?.role || 'exploitant');
      setPoste(data?.poste || data?.role || 'exploitant');
      setProfil(data || null);
    }
    loadRole();
  }, [session]);

  async function handleLogout() {
    await supabase.auth.signOut();
  }

  if (publicToken) {
    return <PublicAttestationForm token={publicToken} />;
  }

  if (session === undefined) {
    return (
      <div className="min-h-screen bg-[#f5f5f5] flex items-center justify-center text-[#d4006e] font-semibold">
        Vérification de la connexion…
      </div>
    );
  }

  if (!session) {
    return <Login />;
  }

  if (role === null) {
    return (
      <div className="min-h-screen bg-[#f5f5f5] flex items-center justify-center text-[#d4006e] font-semibold">
        Vérification des droits d'accès…
      </div>
    );
  }

  return <AppShell userEmail={session.user.email} onLogout={handleLogout} role={role} poste={poste} profil={profil} />;
}

// ===== FORMULAIRE PUBLIC SOUS-TRAITANT =====
// Accessible sans connexion via un lien unique (?token=...).
// Le sous-traitant y renseigne les coordonnées de son chauffeur
// et dépose ses documents (licence, assurance, capacité, K-bis).
function PublicAttestationForm({ token }) {
  const [loading, setLoading] = useState(true);
  const [notFound, setNotFound] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [nomSousTraitant, setNomSousTraitant] = useState('');
  const [form, setForm] = useState({ chauffeurNom: '', chauffeurTel: '' });
  const [files, setFiles] = useState({ licence: null, assurance: null, kbis: null });
  // Le sous-traitant saisit lui-même les dates de validité de chaque
  // document au moment du dépôt : le statut (Valide / Expire bientôt /
  // Expiré) est ensuite calculé automatiquement à partir de ces dates,
  // sans aucune validation humaine supplémentaire.
  const [dates, setDates] = useState({
    licenceDebut: '', licenceFin: '',
    assuranceDebut: '', assuranceFin: '',
    kbisDebut: '', kbisFin: '',
  });

  useEffect(() => {
    async function load() {
      const { data, error } = await supabase.rpc('get_attestation_public', { p_token: token });
      if (error || !data || data.length === 0) {
        setNotFound(true);
        setLoading(false);
        return;
      }
      const row = data[0];
      setNomSousTraitant(row.nom_sous_traitant || '');
      setForm({
        chauffeurNom: row.chauffeur_nom || '',
        chauffeurTel: row.chauffeur_tel || '',
      });
      setLoading(false);
    }
    load();
  }, [token]);

  async function uploadIfPresent(file, docType) {
    if (!file) return null;
    const path = `${token}/${docType}-${Date.now()}-${file.name}`;
    const { error } = await supabase.storage.from('documents-st').upload(path, file);
    if (error) {
      console.error(error);
      return null;
    }
    return path;
  }

  async function handleSubmit() {
    // Si un fichier est déposé pour un document, sa date de fin de
    // validité devient obligatoire (sinon le statut ne pourrait pas être
    // calculé automatiquement).
    if (files.licence && !dates.licenceFin) { alert('Merci de renseigner la date de fin de validité de la licence de transport.'); return; }
    if (files.assurance && !dates.assuranceFin) { alert("Merci de renseigner la date de fin de validité de l'assurance."); return; }
    if (files.kbis && !dates.kbisFin) { alert('Merci de renseigner la date de fin de validité du K-bis.'); return; }

    setSubmitting(true);
    const [licencePath, assurancePath, kbisPath] = await Promise.all([
      uploadIfPresent(files.licence, 'licence'),
      uploadIfPresent(files.assurance, 'assurance'),
      uploadIfPresent(files.kbis, 'kbis'),
    ]);

    const { error } = await supabase.rpc('submit_attestation_public', {
      p_token: token,
      p_chauffeur_nom: form.chauffeurNom || null,
      p_chauffeur_tel: form.chauffeurTel || null,
      p_licence_url: licencePath,
      p_licence_date_debut: dates.licenceDebut || null,
      p_licence_date: dates.licenceFin || null,
      p_assurance_url: assurancePath,
      p_assurance_date_debut: dates.assuranceDebut || null,
      p_assurance_date: dates.assuranceFin || null,
      p_kbis_url: kbisPath,
      p_kbis_date_debut: dates.kbisDebut || null,
      p_kbis_date: dates.kbisFin || null,
    });

    setSubmitting(false);
    if (error) {
      console.error(error);
      alert("Une erreur s'est produite. Merci de réessayer ou de contacter IBK Euro Afrique.");
      return;
    }
    setSubmitted(true);
  }

  if (loading) {
    return (
      <div className="min-h-screen bg-[#f5f5f5] flex items-center justify-center text-[#d4006e] font-semibold">
        Chargement…
      </div>
    );
  }

  if (notFound) {
    return (
      <div className="min-h-screen bg-[#f5f5f5] flex items-center justify-center px-5">
        <div className="bg-white rounded-xl shadow-lg p-8 max-w-md text-center">
          <div className="text-[18px] font-bold text-[#1c2733] mb-2">Lien invalide ou expiré</div>
          <div className="text-[13px] text-[#64748b]">Merci de contacter IBK Euro Afrique pour obtenir un nouveau lien.</div>
        </div>
      </div>
    );
  }

  if (submitted) {
    return (
      <div className="min-h-screen bg-[#f5f5f5] flex items-center justify-center px-5">
        <div className="bg-white rounded-xl shadow-lg p-8 max-w-md text-center">
          <div className="text-[40px] mb-3">✅</div>
          <div className="text-[18px] font-bold text-[#1c2733] mb-2">Merci !</div>
          <div className="text-[13px] text-[#64748b]">Vos informations et documents ont bien été transmis à IBK Euro Afrique. Vous pouvez fermer cette page.</div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#f5f5f5]">
      <div className="h-14 bg-[#0d1b2a] flex items-center px-5 gap-2.5">
        <img src={LOGO_IBK_BASE64} alt="IBK Euro Afrique" className="h-9 w-9 rounded-md object-cover" />
        <span className="text-white font-bold text-[15px]" style={{ fontFamily: "'Space Grotesk', sans-serif" }}>IBK TMS — Espace sous-traitant</span>
      </div>

      <div className="max-w-2xl mx-auto px-5 py-8">
        <div className="bg-white rounded-xl shadow-[0_4px_18px_rgba(13,27,42,0.12)] border border-[#d4006e]/20 p-7">
          <div className="text-[20px] font-bold text-[#1c2733] mb-1">{nomSousTraitant || 'Sous-traitant'}</div>
          <div className="text-[13px] text-[#64748b] mb-6">Merci de renseigner les coordonnées de votre chauffeur et de déposer vos documents à jour. Ces informations sont transmises directement et en toute confidentialité à IBK Euro Afrique.</div>

          <div className="text-[11px] uppercase tracking-wide text-[#64748b] mb-3">Coordonnées du chauffeur</div>
          <FormRow>
            <Field label="Nom du chauffeur">
              <input className="form-input" value={form.chauffeurNom} onChange={(e) => setForm({ ...form, chauffeurNom: e.target.value })} placeholder="Nom et prénom" />
            </Field>
            <Field label="Téléphone">
              <input className="form-input" value={form.chauffeurTel} onChange={(e) => setForm({ ...form, chauffeurTel: e.target.value })} placeholder="06 00 00 00 00" />
            </Field>
          </FormRow>

          <div className="text-[11px] uppercase tracking-wide text-[#64748b] mt-6 mb-3 pt-4 border-t border-gray-100">Documents (PDF ou photo)</div>
          <div className="text-[10px] text-[#64748b] mb-4 -mt-1">💡 Merci d'indiquer la période de validité de chaque document — elle nous permet de suivre automatiquement leur expiration.</div>

          <FileField label="Licence de transport" file={files.licence} onChange={(e) => setFiles({ ...files, licence: e.target.files[0] })} />
          <FormRow>
            <Field label="Licence — Date de début de validité">
              <input type="date" className="form-input" value={dates.licenceDebut} onChange={(e) => setDates({ ...dates, licenceDebut: e.target.value })} />
            </Field>
            <Field label="Licence — Date de fin de validité">
              <input type="date" className="form-input" value={dates.licenceFin} onChange={(e) => setDates({ ...dates, licenceFin: e.target.value })} />
            </Field>
          </FormRow>

          <FileField label="Assurance RC" file={files.assurance} onChange={(e) => setFiles({ ...files, assurance: e.target.files[0] })} />
          <FormRow>
            <Field label="Assurance — Date de début de validité">
              <input type="date" className="form-input" value={dates.assuranceDebut} onChange={(e) => setDates({ ...dates, assuranceDebut: e.target.value })} />
            </Field>
            <Field label="Assurance — Date de fin de validité">
              <input type="date" className="form-input" value={dates.assuranceFin} onChange={(e) => setDates({ ...dates, assuranceFin: e.target.value })} />
            </Field>
          </FormRow>

          <FileField label="K-bis" file={files.kbis} onChange={(e) => setFiles({ ...files, kbis: e.target.files[0] })} />
          <FormRow>
            <Field label="K-bis — Date de début de validité">
              <input type="date" className="form-input" value={dates.kbisDebut} onChange={(e) => setDates({ ...dates, kbisDebut: e.target.value })} />
            </Field>
            <Field label="K-bis — Date de fin de validité">
              <input type="date" className="form-input" value={dates.kbisFin} onChange={(e) => setDates({ ...dates, kbisFin: e.target.value })} />
            </Field>
          </FormRow>

          <button
            onClick={handleSubmit}
            disabled={submitting}
            className="w-full mt-6 px-4 py-3 rounded-md text-[14px] font-semibold bg-[#d4006e] text-white hover:bg-[#b3005d] transition disabled:opacity-60"
          >
            {submitting ? 'Envoi en cours…' : 'Envoyer mes informations et documents'}
          </button>
        </div>
      </div>
    </div>
  );
}
