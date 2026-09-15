# -*- coding: utf-8 -*-
"""Mappa fra le categorie Revit e gli schemi di nomenclatura ESA.

Questo modulo e' dichiarativo: descrive quali schede di classificazione
esistono, quali tabelle del foglio DV alimentano i loro menu, quali blocchi
compongono il nome e da dove si legge il blocco dimensionale. Non importa
nulla dalle API di Revit, cosi' resta leggibile anche a chi non programma
e si puo' correggere senza toccare la logica.

Le tabelle vere stanno in FamilyNaming_data.py, che e' generato dagli .xlsx.

--------------------------------------------------------------------------
IL TYPE MARK

Ha sempre la forma AA-XXX: due lettere, trattino, tre cifre. Cambia da dove
escono le lettere e come si spezzano le cifre, e la chiave 'tm_from' lo dice.

  'g2_digit'   prefisso = codice categoria, cifre = cifra del Group2 + 2 di
               sequenziale. Le otto categorie caricabili con Group1 proprio.
               DR.HN.IN -> DR-101
  'g2_code'    prefisso = codice del Group2, cifre = 3 di sequenziale. Le
               sette categorie caricabili a Group1 condiviso: il codice di
               categoria NON entra nel Type Mark.
               FN.CO.SE -> SE-001
  'g1_digit'   prefisso = codice categoria, cifre = cifra del Group1 + 2 di
               sequenziale. Tutte le schede System e Structural.
               BW.FR.IN -> BW-101

In tutti e tre i casi, con Group1 = OT il Type Mark diventa manuale e prende
un prefisso materiale dalla tabella Tbl_Mat: WD-001, PT-001, MT-001.
Su STRUCTURAL:05_FR con il campo Use compilato il prefisso e' il codice Use:
SK-001, BH-001.

Il sequenziale si cerca sempre PER PREFISSO sull'intero documento, mai per
categoria: con 'g2_code' il codice di categoria non compare nel Type Mark e
lo stesso Group2 ricorre su piu' categorie (SE e' Seating sia su Furniture
sia su Furniture Systems), quindi numerare per categoria produrrebbe due
SE-001 diversi. Dove il prefisso contiene gia' la categoria il risultato e'
identico, quindi la regola resta una sola.

--------------------------------------------------------------------------
I BLOCCHI DIMENSIONALI

  'wxh'        800x2100      larghezza x altezza
  'wxdxh'      600x600x800   larghezza x profondita', altezza solo se serve
  'wxl'        250x500       larghezza x lunghezza
  'wxd_or_d'   400x400 / D500    sezione, tonda -> solo diametro
  'bxh_or_d'   60x40 / D40       profilo, tondo -> solo diametro
  't'          T200          spessore
  'h'          H1100         altezza
  'nfin_t'     2F.125        facce finite . spessore totale
  'free'       HEA200        designazione normalizzata o misure, testo libero
  'none'       -             nessun blocco

--------------------------------------------------------------------------
LE SORGENTI DEL BLOCCO DIMENSIONALE

Ogni misura porta una lista di tentativi, provati in ordine, primo che
risponde vince. Se non risponde nessuno il campo resta vuoto con un
suggerimento a schermo, e l'utente lo compila a mano.

  ('bip',  'DOOR_WIDTH')   parametro di tipo per BuiltInParameter
  ('name', 'Width')        parametro di tipo per nome
  ('prop', 'wall_width')   caso speciale gestito nello script

Tutte le misure si leggono in piedi interni e si convertono in millimetri
interi arrotondati.
"""

# ---------------------------------------------------------------------------
# Tabelle condivise, valide su piu' schede
# ---------------------------------------------------------------------------

AUTHOR_CODE = u"e"

TBL_INPLACE_L = u"LOADABLE:Tbl_InPl"
TBL_MATERIAL = u"SYSTEM:Tbl_Mat"
TBL_SHARED_G1 = u"LOADABLE:Tbl_PR_G1"
TBL_CW_HOSTED = u"LOADABLE:Tbl_CWh"
TBL_ENTRANCE = u"LOADABLE:Tbl_Entr"
TBL_RAIL_TOP = u"SYSTEM:Tbl_RailTop"
TBL_UNDERSIDE = u"SYSTEM:Tbl_Und"
TBL_LN_SHAPE = u"SYSTEM:Tbl_LN_Shape"
TBL_RM_SHAPE = u"SYSTEM:Tbl_RM_Shape"
TBL_NFIN = u"SYSTEM:Tbl_Fin"
TBL_NFIN_01 = u"SYSTEM:Tbl_Fin01"
TBL_ROLE = u"STRUCTURAL:Tbl_Role"
TBL_CUSTOM = u"STRUCTURAL:Tbl_Cst"
TBL_USE = u"STRUCTURAL:Tbl_Use"
TBL_STR_G1 = u"STRUCTURAL:Tbl_STR_G1"

# Codici Group1 che dichiarano una famiglia nidificata: azzerano i campi
# funzionali e forzano Group2 = OT sulle categorie caricabili.
NESTED_G1 = (u"CM", u"SY")

# Codice Group1 che rende il Type Mark interamente manuale.
MANUAL_TM_G1 = u"OT"


# ---------------------------------------------------------------------------
# Sorgenti ricorrenti per il blocco dimensionale
# ---------------------------------------------------------------------------

SRC_DOOR_W = (("bip", "DOOR_WIDTH"), ("bip", "FAMILY_WIDTH_PARAM"), ("name", "Width"))
SRC_DOOR_H = (("bip", "DOOR_HEIGHT"), ("bip", "FAMILY_HEIGHT_PARAM"), ("name", "Height"))
SRC_WIN_W = (("bip", "WINDOW_WIDTH"), ("bip", "FAMILY_WIDTH_PARAM"), ("name", "Width"))
SRC_WIN_H = (("bip", "WINDOW_HEIGHT"), ("bip", "FAMILY_HEIGHT_PARAM"), ("name", "Height"))
SRC_GEN_W = (("bip", "FAMILY_WIDTH_PARAM"), ("name", "Width"), ("name", "Larghezza"))
SRC_GEN_D = (("name", "Depth"), ("name", "Profondita"), ("name", "Profondita'"))
SRC_GEN_H = (("bip", "FAMILY_HEIGHT_PARAM"), ("name", "Height"), ("name", "Altezza"))
SRC_GEN_L = (("name", "Length"), ("name", "Lunghezza"))
SRC_DIA = (("name", "Diameter"), ("name", "Diametro"), ("name", "d"))

# Casi speciali risolti nello script, dove il dato non e' un parametro di tipo
SRC_WALL_T = (("prop", "wall_width"),)
SRC_LAYER_T = (("prop", "compound_width"),)
SRC_RAIL_H = (("prop", "railing_height"), ("name", "Height"))


# ---------------------------------------------------------------------------
# Le trentasei schede
# ---------------------------------------------------------------------------
#
# label        nome della scheda a schermo
# schema       'loadable' genera nome famiglia + nome tipo
#              'system'   genera il solo nome tipo
# cat_table    tabella delle famiglie di sistema o del codice categoria
# g1 / g2      nome tabella, oppure dizionario codice categoria -> tabella
#              quando la scheda ha sottogruppi con liste distinte
# tm_from      vedi la nota sul Type Mark in testa al modulo
# dim          chiave del blocco dimensionale
# dim_src      misura -> lista di tentativi di lettura
# blocks       blocchi opzionali della scheda, nell'ordine in cui compaiono
#              nel nome
# ---------------------------------------------------------------------------

SHEETS = {

    # -- Loadable, Group1 proprio della categoria --------------------------

    u"LOADABLE:01_DR": {
        "label": u"Doors",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_DR_Cat",
        "g1": u"LOADABLE:Tbl_DR_G1",
        "g2": u"LOADABLE:Tbl_DR_G2",
        "tm_from": "g2_digit",
        "dim": "wxh",
        "dim_src": {"w": SRC_DOOR_W, "h": SRC_DOOR_H},
        "blocks": ("leaves", "cw_hosted", "entrance", "rei", "manufacturer",
                   "brand", "description"),
    },
    u"LOADABLE:02_WN": {
        "label": u"Windows",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_WN_Cat",
        "g1": u"LOADABLE:Tbl_WN_G1",
        "g2": u"LOADABLE:Tbl_WN_G2",
        "tm_from": "g2_digit",
        "dim": "wxh",
        "dim_src": {"w": SRC_WIN_W, "h": SRC_WIN_H},
        "blocks": ("leaves", "cw_hosted", "manufacturer", "brand", "description"),
    },
    u"LOADABLE:03_CP": {
        "label": u"Curtain Panels",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_CP_Cat",
        "g1": u"LOADABLE:Tbl_CP_G1",
        "g2": u"LOADABLE:Tbl_CP_G2",
        "tm_from": "g2_digit",
        "dim": "t",
        "dim_src": {"t": SRC_LAYER_T},
        "blocks": ("manufacturer", "brand", "description"),
    },
    u"LOADABLE:04_CM": {
        "label": u"Columns",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_CM_Cat",
        "g1": u"LOADABLE:Tbl_CM_G1",
        "g2": u"LOADABLE:Tbl_CM_G2",
        "tm_from": "g2_digit",
        "dim": "wxd_or_d",
        "dim_src": {"w": SRC_GEN_W, "d": SRC_GEN_D, "dia": SRC_DIA},
        "round_g1": (u"CR",),
        "blocks": ("manufacturer", "brand", "description"),
    },
    u"LOADABLE:11_PK": {
        "label": u"Parking",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_PK_Cat",
        "g1": u"LOADABLE:Tbl_PK_G1",
        "g2": u"LOADABLE:Tbl_PK_G2",
        "tm_from": "g2_digit",
        "dim": "wxl",
        "dim_src": {"w": SRC_GEN_W, "l": SRC_GEN_L},
        "blocks": ("manufacturer", "brand", "description"),
    },
    u"LOADABLE:12_EN": {
        "label": u"Entourage",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_EN_Cat",
        "g1": u"LOADABLE:Tbl_EN_G1",
        "g2": u"LOADABLE:Tbl_EN_G2",
        "tm_from": "g2_digit",
        "dim": "none",
        "dim_src": {},
        "blocks": ("manufacturer", "brand", "description"),
    },
    u"LOADABLE:13_PL": {
        "label": u"Planting",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_PL_Cat",
        "g1": u"LOADABLE:Tbl_PL_G1",
        "g2": u"LOADABLE:Tbl_PL_G2",
        "tm_from": "g2_digit",
        "dim": "h",
        "dim_src": {"h": SRC_GEN_H},
        "blocks": ("manufacturer", "brand", "description"),
    },
    u"LOADABLE:15_MS": {
        "label": u"Mass",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_MS_Cat",
        "g1": u"LOADABLE:Tbl_MS_G1",
        "g2": u"LOADABLE:Tbl_MS_G2",
        "tm_from": "g2_digit",
        "dim": "none",
        "dim_src": {},
        "blocks": ("manufacturer", "brand", "description"),
    },

    # -- Loadable, Group1 condiviso: il Type Mark parte dal Group2 ---------

    u"LOADABLE:05_FN": {
        "label": u"Furniture",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_FN_Cat",
        "g1": TBL_SHARED_G1,
        "g2": u"LOADABLE:Tbl_FN_G2",
        "tm_from": "g2_code",
        "dim": "wxdxh",
        "dim_src": {"w": SRC_GEN_W, "d": SRC_GEN_D, "h": SRC_GEN_H},
        "blocks": ("manufacturer", "brand", "description"),
    },
    u"LOADABLE:06_FS": {
        "label": u"Furniture Systems",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_FS_Cat",
        "g1": TBL_SHARED_G1,
        "g2": u"LOADABLE:Tbl_FS_G2",
        "tm_from": "g2_code",
        "dim": "wxdxh",
        "dim_src": {"w": SRC_GEN_W, "d": SRC_GEN_D, "h": SRC_GEN_H},
        "blocks": ("manufacturer", "brand", "description"),
    },
    u"LOADABLE:07_CK": {
        "label": u"Casework",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_CK_Cat",
        "g1": TBL_SHARED_G1,
        "g2": u"LOADABLE:Tbl_CK_G2",
        "tm_from": "g2_code",
        "dim": "wxdxh",
        "dim_src": {"w": SRC_GEN_W, "d": SRC_GEN_D, "h": SRC_GEN_H},
        "blocks": ("manufacturer", "brand", "description"),
    },
    u"LOADABLE:08_PF": {
        "label": u"Plumbing Fixtures",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_PF_Cat",
        "g1": TBL_SHARED_G1,
        "g2": u"LOADABLE:Tbl_PF_G2",
        "tm_from": "g2_code",
        "dim": "wxdxh",
        "dim_src": {"w": SRC_GEN_W, "d": SRC_GEN_D, "h": SRC_GEN_H},
        "blocks": ("manufacturer", "brand", "description"),
    },
    u"LOADABLE:09_SE": {
        "label": u"Specialty Equipment",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_SE_Cat",
        "g1": TBL_SHARED_G1,
        "g2": u"LOADABLE:Tbl_SE_G2",
        "tm_from": "g2_code",
        "dim": "wxdxh",
        "dim_src": {"w": SRC_GEN_W, "d": SRC_GEN_D, "h": SRC_GEN_H},
        "blocks": ("manufacturer", "brand", "description"),
    },
    u"LOADABLE:10_SI": {
        "label": u"Site",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_SI_Cat",
        "g1": TBL_SHARED_G1,
        "g2": u"LOADABLE:Tbl_SI_G2",
        "tm_from": "g2_code",
        "dim": "wxdxh",
        "dim_src": {"w": SRC_GEN_W, "d": SRC_GEN_D, "h": SRC_GEN_H},
        "blocks": ("manufacturer", "brand", "description"),
    },
    u"LOADABLE:14_GM": {
        "label": u"Generic Models",
        "schema": "loadable",
        "cat_table": u"LOADABLE:Tbl_GM_Cat",
        "g1": TBL_SHARED_G1,
        "g2": u"LOADABLE:Tbl_GM_G2",
        "tm_from": "g2_code",
        "dim": "wxdxh",
        "dim_src": {"w": SRC_GEN_W, "d": SRC_GEN_D, "h": SRC_GEN_H},
        "blocks": ("manufacturer", "brand", "description"),
    },

    # -- System: generano il solo nome del tipo ----------------------------

    u"SYSTEM:01_CL": {
        "label": u"Ceilings",
        "schema": "system",
        "cat_table": u"SYSTEM:Tbl_CL_Cat",
        "g1": u"SYSTEM:Tbl_CL_G1",
        "g2": u"SYSTEM:Tbl_CL_G2",
        "tm_from": "g1_digit",
        "dim": "nfin_t",
        "dim_src": {"t": SRC_LAYER_T},
        "nfin_table": TBL_NFIN_01,
        "blocks": ("rei", "wi", "description"),
        "wi_allowed_g2": (u"IN",),
    },
    u"SYSTEM:02_FL": {
        "label": u"Floors",
        "schema": "system",
        "cat_table": u"SYSTEM:Tbl_FL_Cat",
        "g1": u"SYSTEM:Tbl_FL_G1",
        "g2": u"SYSTEM:Tbl_FL_G2",
        "tm_from": "g1_digit",
        "dim": "nfin_t",
        "dim_src": {"t": SRC_LAYER_T},
        "nfin_table": TBL_NFIN,
        "blocks": ("rei", "wi", "description"),
        "wi_allowed_g2": (u"IN",),
    },
    u"SYSTEM:03_RF": {
        "label": u"Roofs",
        "schema": "system",
        "cat_table": u"SYSTEM:Tbl_RF_Cat",
        "g1": {u"BR": u"SYSTEM:Tbl_BR_G1", u"SZ": u"SYSTEM:Tbl_SZ_G1"},
        "g2": {u"BR": u"SYSTEM:Tbl_BR_G2", u"SZ": u"SYSTEM:Tbl_SZ_G2"},
        "tm_from": "g1_digit",
        "dim": "nfin_t",
        "dim_src": {"t": SRC_LAYER_T},
        "nfin_table": TBL_NFIN,
        "blocks": ("rei", "description"),
    },
    u"SYSTEM:04_RL": {
        "label": u"Railings",
        "schema": "system",
        "cat_table": u"SYSTEM:Tbl_RL_Cat",
        "g1": u"SYSTEM:Tbl_RL_G1",
        "g2": u"SYSTEM:Tbl_RL_G2",
        "tm_from": "g1_digit",
        "dim": "h",
        "dim_src": {"h": SRC_RAIL_H},
        "blocks": ("rail_top", "description"),
    },
    u"SYSTEM:05_RN": {
        "label": u"Stair Runs",
        "schema": "system",
        "cat_table": u"SYSTEM:Tbl_RN_Cat",
        "g1": u"SYSTEM:Tbl_RN_G1",
        "g2": u"SYSTEM:Tbl_RN_G2",
        "tm_from": "g1_digit",
        "dim": "t",
        "dim_src": {"t": SRC_LAYER_T},
        "blocks": ("underside", "description"),
        "underside_cat": (u"MR",),
    },
    u"SYSTEM:06_LN": {
        "label": u"Stair Landings",
        "schema": "system",
        "cat_table": u"SYSTEM:Tbl_LN_Cat",
        "g1": u"SYSTEM:Tbl_LN_G1",
        "g2": u"SYSTEM:Tbl_LN_G2",
        "tm_from": "g1_digit",
        "dim": "t",
        "dim_src": {"t": SRC_LAYER_T},
        "blocks": ("ln_shape", "description"),
    },
    u"SYSTEM:07_ST": {
        "label": u"Stairs",
        "schema": "system",
        "cat_table": u"SYSTEM:Tbl_ST_Cat",
        "g1": {u"AS": u"SYSTEM:Tbl_AS_G1", u"CS": u"SYSTEM:Tbl_CS_G1",
               u"RS": u"SYSTEM:Tbl_RS_G1"},
        "g2": u"SYSTEM:Tbl_ST_G2",
        "tm_from": "g1_digit",
        "dim": "none",
        "dim_src": {},
        "blocks": ("description",),
    },
    u"SYSTEM:08_RM": {
        "label": u"Ramps",
        "schema": "system",
        "cat_table": u"SYSTEM:Tbl_RM_Cat",
        "g1": u"SYSTEM:Tbl_RM_G1",
        "g2": u"SYSTEM:Tbl_RM_G2",
        "tm_from": "g1_digit",
        "dim": "t",
        "dim_src": {"t": SRC_LAYER_T},
        "blocks": ("rm_shape", "description"),
    },
    u"SYSTEM:09_WL": {
        "label": u"Walls",
        "schema": "system",
        "cat_table": u"SYSTEM:Tbl_WL_Cat",
        "g1": {u"BW": u"SYSTEM:Tbl_BW_G1", u"CW": u"SYSTEM:Tbl_CW_G1"},
        "g2": {u"BW": u"SYSTEM:Tbl_BW_G2", u"CW": u"SYSTEM:Tbl_CW_G2"},
        "tm_from": "g1_digit",
        "dim": "nfin_t",
        "dim_src": {"t": SRC_WALL_T},
        "nfin_table": TBL_NFIN,
        "blocks": ("rei", "wi", "description"),
        "wi_allowed_g2": (u"IN", u"SE", u"LN"),
    },
    u"SYSTEM:10_CN": {
        "label": u"Curtain Wall Mullions",
        "schema": "system",
        "cat_table": u"SYSTEM:Tbl_CN_Cat",
        "g1": u"SYSTEM:Tbl_CN_G1",
        "g2": u"SYSTEM:Tbl_CN_G2",
        "tm_from": "g1_digit",
        "dim": "bxh_or_d",
        "dim_src": {"b": SRC_GEN_W, "h": SRC_GEN_D, "dia": SRC_DIA},
        "round_g1": (u"CR",),
        "blocks": ("description",),
    },
    u"SYSTEM:11_TR": {
        "label": u"Top Rails",
        "schema": "system",
        "cat_table": u"SYSTEM:Tbl_TR_Cat",
        "g1": u"SYSTEM:Tbl_TR_G1",
        "g2": u"SYSTEM:Tbl_TR_G2",
        "tm_from": "g1_digit",
        "dim": "bxh_or_d",
        "dim_src": {"b": SRC_GEN_W, "h": SRC_GEN_H, "dia": SRC_DIA},
        "round_g1": (u"CR",),
        "blocks": ("description",),
    },
    u"SYSTEM:12_HR": {
        "label": u"Handrails",
        "schema": "system",
        "cat_table": u"SYSTEM:Tbl_HR_Cat",
        # gli Handrails non hanno una tabella Group1 propria: usano quella
        # dei Top Rails, una sola lista di profili da mantenere
        "g1": u"SYSTEM:Tbl_TR_G1",
        "g2": u"SYSTEM:Tbl_HR_G2",
        "tm_from": "g1_digit",
        "dim": "bxh_or_d",
        "dim_src": {"b": SRC_GEN_W, "h": SRC_GEN_H, "dia": SRC_DIA},
        "round_g1": (u"CR",),
        "blocks": ("description",),
    },
    u"SYSTEM:13_CP": {
        "label": u"Curtain Panels (system)",
        "schema": "system",
        "cat_table": u"SYSTEM:Tbl_CP_Cat",
        "g1": u"SYSTEM:Tbl_CP_G1",
        "g2": u"SYSTEM:Tbl_CP_G2",
        "tm_from": "g1_digit",
        "dim": "t",
        "dim_src": {"t": SRC_LAYER_T},
        # il pannello vuoto non ha spessore, il blocco sparisce
        "dim_optional_g1": (u"EM",),
        "blocks": ("description",),
    },
    u"SYSTEM:14_TS": {
        "label": u"Toposolid",
        "schema": "system",
        "cat_table": u"SYSTEM:Tbl_TS_Cat",
        "g1": u"SYSTEM:Tbl_TS_G1",
        "g2": u"SYSTEM:Tbl_TS_G2",
        "tm_from": "g1_digit",
        "dim": "t",
        "dim_src": {"t": SRC_LAYER_T},
        "blocks": ("description",),
        "min_revit": 2024,
    },

    # -- Structural, famiglie di sistema: nome del tipo --------------------

    u"STRUCTURAL:01_FS": {
        "label": u"Foundation Slabs",
        "schema": "system",
        "cat_table": u"STRUCTURAL:Tbl_FS_Cat",
        "g1": TBL_STR_G1,
        "g2": u"STRUCTURAL:Tbl_FS_G2",
        "tm_from": "g1_digit",
        "dim": "t",
        "dim_src": {"t": SRC_LAYER_T},
        "blocks": ("description",),
    },
    u"STRUCTURAL:02_FW": {
        "label": u"Wall Foundations",
        "schema": "system",
        "cat_table": u"STRUCTURAL:Tbl_FW_Cat",
        "g1": TBL_STR_G1,
        "g2": u"STRUCTURAL:Tbl_FW_G2",
        "tm_from": "g1_digit",
        "dim": "wxt",
        "dim_src": {"w": SRC_GEN_W, "t": SRC_LAYER_T},
        "blocks": ("description",),
    },

    # -- Structural, famiglie caricabili: nome famiglia + nome tipo --------

    u"STRUCTURAL:03_FI": {
        "label": u"Isolated Foundations",
        "schema": "loadable",
        "cat_table": u"STRUCTURAL:Tbl_FI_Cat",
        "g1": TBL_STR_G1,
        "g2": u"STRUCTURAL:Tbl_FI_G2",
        "tm_from": "g1_digit",
        "dim": "free",
        "dim_src": {},
        "dim_hint": u"e.g. 1200x1200x400",
        "blocks": ("role", "custom", "manufacturer", "description"),
    },
    u"STRUCTURAL:04_SC": {
        "label": u"Structural Columns",
        "schema": "loadable",
        "cat_table": u"STRUCTURAL:Tbl_SC_Cat",
        "g1": TBL_STR_G1,
        "g2": u"STRUCTURAL:Tbl_SC_G2",
        "tm_from": "g1_digit",
        "dim": "free",
        "dim_src": {},
        "dim_hint": u"e.g. HEA200, 400x400, D500",
        "blocks": ("role", "custom", "manufacturer", "description"),
    },
    u"STRUCTURAL:05_FR": {
        "label": u"Structural Framing",
        "schema": "loadable",
        "cat_table": u"STRUCTURAL:Tbl_FR_Cat",
        "g1": TBL_STR_G1,
        "g2": u"STRUCTURAL:Tbl_FR_G2",
        "tm_from": "g1_digit",
        "dim": "free",
        "dim_src": {},
        "dim_hint": u"e.g. IPE400, UPN160, 300x600",
        "blocks": ("use", "role", "custom", "manufacturer", "description"),
    },
    u"STRUCTURAL:06_TR": {
        "label": u"Trusses",
        "schema": "loadable",
        "cat_table": u"STRUCTURAL:Tbl_TR_Cat",
        "g1": TBL_STR_G1,
        "g2": u"STRUCTURAL:Tbl_TR_G2",
        "tm_from": "g1_digit",
        "dim": "h",
        "dim_src": {"h": SRC_GEN_H},
        "blocks": ("role", "custom", "manufacturer", "description"),
    },
    u"STRUCTURAL:07_CN": {
        "label": u"Connections",
        "schema": "loadable",
        "cat_table": u"STRUCTURAL:Tbl_CN_Cat",
        "g1": TBL_STR_G1,
        "g2": u"STRUCTURAL:Tbl_CN_G2",
        "tm_from": "g1_digit",
        "dim": "free",
        "dim_src": {},
        "dim_hint": u"e.g. 400x400x20, HPKM20",
        "blocks": ("role", "custom", "manufacturer", "description"),
    },
}


# ---------------------------------------------------------------------------
# Categoria Revit -> scheda
# ---------------------------------------------------------------------------
#
# Il valore e' una tupla di schede candidate. Una sola candidata significa
# nessuna ambiguita'. Piu' candidate significa che la categoria Revit da sola
# non basta e va disambiguata leggendo il tipo: se ne occupa resolve_sheet()
# nello script, che distingue le famiglie caricabili (FamilySymbol) da quelle
# di sistema e, dentro le fondazioni, la platea dal muro dal plinto.
#
# Le categorie che non compaiono qui non sono mappate: il tool avvisa e si
# ferma, come da specifica.
# ---------------------------------------------------------------------------

CATEGORY_MAP = {
    # architettonico caricabile
    "OST_Doors": (u"LOADABLE:01_DR",),
    "OST_Windows": (u"LOADABLE:02_WN",),
    "OST_Columns": (u"LOADABLE:04_CM",),
    "OST_Furniture": (u"LOADABLE:05_FN",),
    "OST_FurnitureSystems": (u"LOADABLE:06_FS",),
    "OST_Casework": (u"LOADABLE:07_CK",),
    "OST_PlumbingFixtures": (u"LOADABLE:08_PF",),
    "OST_SpecialityEquipment": (u"LOADABLE:09_SE",),
    "OST_Site": (u"LOADABLE:10_SI",),
    "OST_Parking": (u"LOADABLE:11_PK",),
    "OST_Entourage": (u"LOADABLE:12_EN",),
    "OST_Planting": (u"LOADABLE:13_PL",),
    "OST_GenericModel": (u"LOADABLE:14_GM",),
    "OST_Mass": (u"LOADABLE:15_MS",),

    # architettonico di sistema
    "OST_Ceilings": (u"SYSTEM:01_CL",),
    "OST_Floors": (u"SYSTEM:02_FL",),
    "OST_EdgeSlab": (u"SYSTEM:02_FL",),          # Slab Edge, codice SG
    "OST_Roofs": (u"SYSTEM:03_RF",),
    "OST_StairsRailing": (u"SYSTEM:04_RL",),
    "OST_Railings": (u"SYSTEM:04_RL",),
    "OST_RailingSystem": (u"SYSTEM:04_RL",),
    "OST_StairsRuns": (u"SYSTEM:05_RN",),
    "OST_StairsLandings": (u"SYSTEM:06_LN",),
    "OST_Stairs": (u"SYSTEM:07_ST",),
    "OST_Ramps": (u"SYSTEM:08_RM",),
    "OST_Walls": (u"SYSTEM:09_WL",),
    "OST_StackedWalls": (u"SYSTEM:09_WL",),
    "OST_CurtainWallMullions": (u"SYSTEM:10_CN",),
    "OST_RailingTopRail": (u"SYSTEM:11_TR",),
    "OST_RailingSystemTopRail": (u"SYSTEM:11_TR",),
    "OST_RailingHandRail": (u"SYSTEM:12_HR",),
    "OST_RailingSystemHandRail": (u"SYSTEM:12_HR",),
    "OST_Toposolid": (u"SYSTEM:14_TS",),         # solo da Revit 2024

    # strutturale
    "OST_StructuralColumns": (u"STRUCTURAL:04_SC",),
    "OST_StructuralFraming": (u"STRUCTURAL:05_FR",),
    "OST_StructuralTruss": (u"STRUCTURAL:06_TR",),
    "OST_StructConnections": (u"STRUCTURAL:07_CN",),
    "OST_StructuralStiffener": (u"STRUCTURAL:07_CN",),

    # L'unica ambiguita' rimasta. Platea, trave rovescia e plinto stanno tutti
    # in OST_StructuralFoundation: le prime due sono famiglie di sistema, il
    # plinto e' caricabile, e resolve_sheet() le distingue sulla classe.
    "OST_StructuralFoundation": (u"STRUCTURAL:01_FS", u"STRUCTURAL:02_FW",
                                 u"STRUCTURAL:03_FI"),
    # pannello di sistema oppure famiglia caricabile annidata nella griglia
    "OST_CurtainWallPanels": (u"SYSTEM:13_CP", u"LOADABLE:03_CP"),
}


# Categorie Revit che fissano da sole il codice della famiglia di sistema,
# perche' in Revit sono categorie distinte anche se la scheda le tiene
# insieme. Senza questa tabella un bordo di solaio finirebbe nominato come
# un solaio.
FORCED_CAT_CODE = {
    "OST_EdgeSlab": u"SG",
    "OST_Floors": u"FL",
}


# ---------------------------------------------------------------------------
# Quali note mostrare
# ---------------------------------------------------------------------------
#
# I fogli generatore portano in coda 559 righe di note, divise in due blocchi:
# le NOTE della categoria, che dicono come classificare, e le REGOLE DI
# COMPILAZIONE, che spiegano campo per campo cosa scrivere in ogni colonna.
#
# Il secondo blocco e' comodo dentro l'Excel, dove non c'e' nessuno a
# spiegarti le colonne, ma nel tool e' ridondante: ogni campo porta gia' il
# suo suggerimento accanto, i due gruppi hanno la descrizione viva sotto la
# tendina, e categoria e misure sono mostrate come valori calcolati.
#
# Il filtro sta qui e non in FamilyNaming_data.py perche' quello e' generato:
# toglierle di la' significherebbe vederle tornare alla prossima
# rigenerazione dagli Excel.
# ---------------------------------------------------------------------------

# Le istruzioni campo per campo. False le nasconde tutte in blocco.
NOTES_SHOW_FIELD_RULES = False

# Note di categoria da non mostrare, per intestazione. Valgono su tutte le
# schede in cui quella intestazione compare.
NOTES_HIDDEN = (
    u"Building Pad",
    u"Categoria di ultima istanza",
    u"Contesto",
    u"Elementi in place",
    u"Family role è quasi sempre Component",
    u"Fissaggio",
    u"Fondazioni di sistema e caricabili nello stesso file",
    u"Gli irrigidimenti stanno qui",
    u"Gli usi non strutturali della categoria",
    u"Il ruolo strutturale resta fuori",
    u"Il template della famiglia",
    u"Impianti",
    u"La famiglia è il tipo di fondazione, il tipo è la misura",
    u"La trave di fondazione",
    u"Le due dimensioni",
    u"Manufacturer",
    u"Manufacturer e Brand",
    u"NOTE — CASEWORK",
    u"NOTE — CEILINGS",
    u"NOTE — COLUMNS",
    u"NOTE — CONNECTIONS",
    u"NOTE — CURTAIN PANELS",
    u"NOTE — CURTAIN WALL MULLIONS",
    u"NOTE — DOORS",
    u"NOTE — ENTOURAGE",
    u"NOTE — FLOORS",
    u"NOTE — FOUNDATION SLABS",
    u"NOTE — FURNITURE",
    u"NOTE — FURNITURE SYSTEMS",
    u"NOTE — GENERIC MODELS",
    u"NOTE — HANDRAILS",
    u"NOTE — ISOLATED FOUNDATIONS",
    u"NOTE — LANDINGS",
    u"NOTE — MASS",
    u"NOTE — PARKING",
    u"NOTE — PLANTING",
    u"NOTE — PLUMBING FIXTURES",
    u"NOTE — RAILINGS",
    u"NOTE — RAMPS",
    u"NOTE — ROOFS",
    u"NOTE — RUNS",
    u"NOTE — SITE",
    u"NOTE — SPECIALTY EQUIPMENT",
    u"NOTE — STAIRS",
    u"NOTE — STRUCTURAL COLUMNS",
    u"NOTE — STRUCTURAL FRAMING",
    u"NOTE — TOP RAILS",
    u"NOTE — TOPOSOLID",
    u"NOTE — TRUSSES",
    u"NOTE — WALL FOUNDATIONS",
    u"NOTE — WALLS",
    u"NOTE — WINDOWS",
    u"Niente altezza",
    u"Niente campo Use",
    u"Niente numero di finiture",
    u"Niente resistenza al fuoco",
    u"Non sono le rampe di scala",
    u"Pali e plinti su pali",
    u"Pannelli vuoti",
    u"Pendenza",
    u"Peso del modello",
    u"Porte in facciata continua",
    u"Site o famiglie di sistema",
    u"SlabEdge",
    u"Solo Compound Ceiling",
    u"Solo pannelli caricabili",
    u"Sovrapposizione con i Floors",
    u"Terreno e pavimentazioni",
)

# Eccezioni: la stessa intestazione porta testi diversi su schede diverse e
# va nascosta solo dove il testo non aggiunge nulla. Oggi ce n'e' una sola:
# la nota Group1 spiega la tabella della provenienza, e serve solo sulle
# sette categorie che quella tabella la usano davvero.
NOTES_HIDDEN_ON = {
    u"Group1": (
        u"LOADABLE:01_DR",
        u"LOADABLE:02_WN",
        u"LOADABLE:03_CP",
        u"LOADABLE:04_CM",
        u"LOADABLE:11_PK",
        u"LOADABLE:12_EN",
        u"LOADABLE:13_PL",
        u"LOADABLE:15_MS",
    ),
}
