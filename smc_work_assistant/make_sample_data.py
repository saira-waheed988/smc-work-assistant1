"""Creates synthetic sample files (no real SMC data)."""
import pandas as pd

old = pd.DataFrame({
    "Room ID": ["GF-01", "GF-02", "GF-03", "GF-04", "FF-01", "FF-02", "FF-03", "FF-04"],
    "Space Name": ["Lobby", "Reception", "Server Room", "Pantry", "Office A", "Office B", "Meeting", "Store"],
    "Floor": ["Ground"] * 4 + ["First"] * 4,
    "Area (m2)": [120.5, 45.0, 30.25, 18.0, 85.0, 85.0, 40.0, 12.0],
})
new = old.copy()
new.loc[new["Room ID"] == "GF-02", "Area (m2)"] = 48.5          # modified
new.loc[new["Room ID"] == "FF-03", "Space Name"] = "Board Room"  # modified
new = new[new["Room ID"] != "FF-04"]                             # removed
new = pd.concat([new, pd.DataFrame({"Room ID": ["FF-05", "SF-01"], "Space Name": ["Office C", "Control Room"],
                                    "Floor": ["First", "Second"], "Area (m2)": [60.0, 55.0]})])
old.to_excel("sample_data/area_previous.xlsx", index=False)
new.to_excel("sample_data/area_updated.xlsx", index=False)
print("Sample files created.")
