# -*- coding: utf-8 -*-
__title__   = "Super\nTo Shared"
__doc__     = """Version = 1.0
Date    = 08.05.2025
________________________________________________________________
Trasferisce i valori dei parametri selezionati dalla famiglia host alle famiglie nidificate condivise.
________________________________________________________________
Author(s): 
Andrea Patti
"""

# REFERENCES
from pyrevit import revit, script
from rpw.ui.forms import FlexForm, Label, TextBox, Button
from Autodesk.Revit.DB import Transaction, StorageType, FamilyInstance, BuiltInParameter
import Autodesk.Revit.DB as DB
import traceback

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

# Interfaccia utente
from rpw.ui.forms import FlexForm, Label, TextBox, Button, ComboBox, Separator

components = [
    Label("Parametri da copiare (separati da ;)"),
    TextBox("parametri", Width=400),
    Separator(),
    Label("Modalità di selezione:"),
    ComboBox("modalita", ["Elementi selezionati", "Intero progetto"]),
    Separator(),
    Button("Esegui")
]

form = FlexForm("Host to Nested Shared", components)
form.show()

# Esci se il form è stato annullato
if not form.values:
    script.exit()

# Parametri da copiare
parametri_da_copiare = [p.strip() for p in form.values["parametri"].split(";") if p.strip()]

if not parametri_da_copiare:
    script.exit(message="⚠️ Nessun parametro inserito.")

# Ottieni gli elementi in base alla modalità selezionata
modalita = form.values["modalita"]
elementi_host = []

# Funzione per filtrare solo le FamilyInstance
def filtra_family_instances(elementi):
    return [elem for elem in elementi if isinstance(elem, FamilyInstance)]

if modalita == "Elementi selezionati":
    # Elementi selezionati
    selection_ids = uidoc.Selection.GetElementIds()
    print("Elementi selezionati: {}".format(len(selection_ids)))
    if not selection_ids:
        script.exit(message="⚠️ Nessun elemento selezionato.")
    elementi_host = filtra_family_instances([doc.GetElement(id) for id in selection_ids])

else:  # "Intero progetto"
    # Tutti gli elementi del progetto
    try:
        output.print_md("Raccolta di tutti gli elementi del progetto...")
        elementi = DB.FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        elementi_host = filtra_family_instances(elementi)
        print("Totale elementi trovati: {}".format(len(elementi_host)))
    except Exception as e:
        output.print_md("Errore durante la raccolta degli elementi: {}".format(e))
        output.print_md(traceback.format_exc())
        script.exit(message="⚠️ Errore nel recupero degli elementi del progetto.")

print("Elementi host validi: {}".format(len(elementi_host)))
if not elementi_host:
    script.exit(message="⚠️ Nessuna famiglia host valida trovata.")

# Verifica se la famiglia è condivisa
def is_shared(instance):
    try:
        # Soluzione alternativa per verificare se una famiglia è condivisa
        family = instance.Symbol.Family
        # Prova ad accedere alla proprietà IsShared in diversi modi
        try:
            return family.get_Parameter(BuiltInParameter.FAMILY_SHARED).AsInteger() == 1
        except:
            pass
            
        try:
            shared_param = family.Parameters.get_Item("Shared")
            if shared_param:
                return shared_param.AsInteger() == 1
        except:
            pass
            
        # Ottieni direttamente dalla proprietà Family.IsShared se possibile
        try:
            return family.IsShared
        except:
            pass
            
        # Se nessuno dei metodi funziona, controlla se la famiglia ha un'identificazione condivisa
        try:
            return family.GetFamilySymbolIds().Count > 0
        except:
            pass
            
        return False
    except Exception as e:
        output.print_md("Errore verifica famiglia condivisa: {}".format(e))
        return False

# Estrai le nidificate condivise
def get_nested_shared(host):
    nested = []
    try:
        for sub_id in host.GetSubComponentIds():
            sub_el = doc.GetElement(sub_id)
            if isinstance(sub_el, FamilyInstance) and is_shared(sub_el):
                nested.append(sub_el)
                # Correzione: uso di sub_el invece di nested
                # print("Famiglia nidificata condivisa: {}".format(sub_el.Name))
    except Exception as e:
        output.print_md("Errore nel parsing nidificato: {}".format(e))
    return nested

# Copia valore parametro
def copia_valore(source, dest, nome_param):
    try:
        p_source = next((p for p in source.Parameters if p.Definition.Name == nome_param), None)
        p_dest = next((p for p in dest.Parameters if p.Definition.Name == nome_param and not p.IsReadOnly), None)
        if not p_source or not p_dest:
            return False

        if p_source.StorageType == StorageType.String:
            p_dest.Set(p_source.AsString() or "")
        elif p_source.StorageType == StorageType.Integer:
            p_dest.Set(p_source.AsInteger())
        elif p_source.StorageType == StorageType.Double:
            p_dest.Set(p_source.AsDouble())
        elif p_source.StorageType == StorageType.ElementId:
            elid = p_source.AsElementId()
            if elid and elid.IntegerValue != -1:
                p_dest.Set(elid)
        return True
    except Exception as e:
        output.print_md("Errore copia '{}': {}".format(nome_param, e))
        output.print_md(traceback.format_exc())
        return False

# Transazione
t = Transaction(doc, "Copia Parametri Host -> Nidificati")
t.Start()

count_host = 0
count_nested = 0
count_parametri = 0

for host in elementi_host:
    nested_list = get_nested_shared(host)
    # print("Famiglia host: {}".format(host.Name))  # Correzione: stampa il nome della famiglia host
    if nested_list:
        count_host += 1
    for nested in nested_list:
        count_nested += 1
        for nome_param in parametri_da_copiare:
            if copia_valore(host, nested, nome_param):
                count_parametri += 1

t.Commit()

# Risultati
output.print_md("# ✅ Operazione Completata")
output.print_md("- Modalità utilizzata: **{}**".format(modalita))
output.print_md("- Famiglie Host elaborate: **{}**".format(count_host))
output.print_md("- Nidificate condivise modificate: **{}**".format(count_nested))
output.print_md("- Parametri copiati: **{}**".format(count_parametri))