import sys
sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2024 SP7\Python\3.10")
import powerfactory as pf

app = pf.GetApplication()
if app is None:
    raise RuntimeError("PowerFactory nicht erreichbar")

project_name = "Nine-bus System(2)"
app.ActivateProject(project_name)

project = app.GetActiveProject()
if project is None:
    raise RuntimeError("Projekt nicht aktiv")

studycase = app.GetActiveStudyCase()
if studycase is None:
    raise RuntimeError("Kein aktiver Berechnungsfall")

print("Study Case:", studycase.loc_name)

def generate_tokens(name):
    name = name.lower()
    tokens = set()

    # Grundform
    tokens.add(name)

    # Leerzeichen / Zahlen entfernen
    tokens.add(name.replace(" ", ""))

    # Übersetzungen
    name = name.replace("load", "last")
    tokens.add(name)
    tokens.add(name.replace(" ", ""))

    return tokens

def build_load_catalog(project):
    catalog = []

    for load in project.GetContents("*.ElmLod", 1):
        entry = {
            "pf_object": load,
            "loc_name": load.loc_name,
            "full_name": load.GetFullName(),
            "tokens": generate_tokens(load.loc_name)
        }
        catalog.append(entry)

    return catalog

def resolve_load(user_input, catalog):
    key = user_input.lower().replace(" ", "")
    for entry in catalog:
        if key in entry["tokens"]:
            return entry["pf_object"]

    return None


catalog = build_load_catalog(project)





# --- LLM-Anweisung (hier manuell simuliert) ---
instruction = {
    "action": "change_load",
    "load_name": "Load A",
    "delta_p_mw": -5
}


resolved_load = resolve_load(instruction["load_name"], catalog)




# --- Interpreter ---

def apply_llm_instruction(instruction, resolved_load):
    if instruction["action"] == "change_load":

        if resolved_load is None:
           raise ValueError(
            f"Keine passende Last für '{instruction['load_name']}' gefunden"
           )
    

        p_old = resolved_load.GetAttribute("plini")
        p_new = p_old + instruction["delta_p_mw"]
        resolved_load.SetAttribute("plini", p_new)

        print(f"{resolved_load.loc_name}: {p_old} → {p_new} MW")


ldf_list = studycase.GetContents("*.ComLdf")

if not ldf_list:
    ldf = studycase.CreateObject("ComLdf", "LoadFlow")
else:
    ldf = ldf_list[0]


# --- Ausführen ---
ldf.Execute()                       # vorher

# Alle Knoten (Busse) im Projekt holen
buses = project.GetContents("*.ElmTerm", 1)

u_before = {}

for bus in buses: 
    name = bus.loc_name
    u = bus.GetAttribute("m:u")
    u_before[name] = u



apply_llm_instruction(instruction, resolved_load)

ldf.Execute()                       # nachher

u_after = {}

for bus in buses: 
    name = bus.loc_name
    u = bus.GetAttribute("m:u")
    u_after[name] = u

for name in u_before:
    delta = u_after[name] - u_before[name]
    print(f"{name:20s}: {delta:+.5f}")
