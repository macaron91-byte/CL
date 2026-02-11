# -*- coding: utf-8 -*-
"""
Dashboard Analyse Moteur - Version avec Interface Graphique
"""

import pandas as pd
import numpy as np
import os
import re
from datetime import datetime
import plotly.graph_objects as go
import glob
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

def detect_nom_colonnes(df, max_lignes=5):
    for i in range(min(max_lignes, len(df))):
        ligne = df.iloc[i]
        n_text = sum(1 for x in ligne if isinstance(x, str) and x.strip() != "")
        if n_text / len(ligne) >= 0.5:
            return i
    return 0

def nettoyer_valeur(x):
    if pd.isna(x):
        return pd.NA
    if isinstance(x, (int, float)):
        return float(x)
    if isinstance(x, str):
        x = x.strip()
        if x == '' or x.lower() in ['nan', 'n/a', '-', '#n/a', 'null']:
            return pd.NA
        x = x.replace(',', '.')
        x = x.replace(' ', '')
        x = re.sub(r'[^0-9.\-+eE]', '', x)
        if x == '' or x == '.' or x == '-' or x == '+':
            return pd.NA
        try:
            val = float(x)
            if pd.isna(val) or val == float('inf') or val == float('-inf'):
                return pd.NA
            return val
        except:
            return pd.NA
    return pd.NA

def extraire_regime(nom_fichier):
    """Extrait le régime moteur du nom de fichier (formats: 1800trmin, 1800rpm, 1800tr/min, 1800 rpm)"""
    nom_lower = nom_fichier.lower()
    
    patterns = [
        r'(\d{3,4})\s*tr/?min',
        r'(\d{3,4})\s*rpm',
        r'(\d{3,4})\s*tr\b',
        r'(\d{3,4})\s*t/min',
        r'(\d{3,4})\s*_rpm',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, nom_lower)
        if match:
            return int(match.group(1))
    
    return None

def analyser_fichiers_liste(fichiers_liste, periode_secondes=60):
    moyennes_fichiers = []
    colonnes_finales = []
    unites_finales = []
    
    print(f"\nAnalyse de {len(fichiers_liste)} fichiers...")
    print(f"Période de moyennage: {periode_secondes} secondes")
    
    for fichier in fichiers_liste:
        print(f"\nTraitement: {os.path.basename(fichier)}")
        try:
            df_full = pd.read_excel(fichier, sheet_name=0, header=None)
        except Exception as e:
            print(f"  Erreur: {e}")
            continue
        
        if len(df_full) < 2:
            continue
        
        idx_nom_col = detect_nom_colonnes(df_full)
        noms_colonnes = df_full.iloc[idx_nom_col].astype(str).tolist()
        
        if idx_nom_col + 1 < len(df_full):
            unite_colonnes = df_full.iloc[idx_nom_col + 1].astype(str).tolist()
            debut_data = idx_nom_col + 2
        else:
            unite_colonnes = [''] * len(noms_colonnes)
            debut_data = idx_nom_col + 1
        
        seen = {}
        noms_uniques = []
        for col in noms_colonnes:
            if col not in seen:
                seen[col] = 1
                noms_uniques.append(col)
            else:
                seen[col] += 1
                noms_uniques.append(f"{col}_{seen[col]}")
        
        for i, col in enumerate(noms_uniques):
            if col not in colonnes_finales:
                colonnes_finales.append(col)
                unites_finales.append(unite_colonnes[i] if i < len(unite_colonnes) else '')
        
        df_data = df_full.iloc[debut_data:].copy()
        df_data.columns = noms_uniques
        
        # Filtrage temporel sur la colonne Heure
        df_filtre = df_data.copy()
        if 'Heure' in df_data.columns:
            def heure_vers_secondes(h):
                if isinstance(h, pd.Timestamp):
                    return h.hour * 3600 + h.minute * 60 + h.second
                elif isinstance(h, str):
                    try:
                        parts = h.split(':')
                        return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
                    except:
                        return None
                return None
            
            df_data['secondes'] = df_data['Heure'].apply(heure_vers_secondes)
            
            if df_data['secondes'].notna().sum() > 0:
                temps_max = df_data['secondes'].max()
                temps_min_filtre = temps_max - periode_secondes
                df_filtre = df_data[df_data['secondes'] >= temps_min_filtre]
                print(f"   Filtrage temporel: {len(df_filtre)}/{len(df_data)} lignes (dernières {periode_secondes}s)")
        
        # Nettoyage des données filtrées
        colonnes_nettoyees = {}
        for col in df_filtre.columns:
            serie_nettoyee = df_filtre[col].apply(nettoyer_valeur)
            colonnes_nettoyees[col] = pd.to_numeric(serie_nettoyee, errors='coerce')
        
        df_numerique = pd.DataFrame(colonnes_nettoyees)
        
        colonnes_numeriques = [col for col in df_numerique.columns 
                               if pd.api.types.is_numeric_dtype(df_numerique[col]) 
                               and df_numerique[col].notna().sum() > 0]
        
        if len(colonnes_numeriques) == 0:
            continue
        
        moyennes = df_numerique[colonnes_numeriques].mean(skipna=True).to_frame().T
        
        # Calcul écart-type pour vérifier la stabilité
        ecarts_types = df_numerique[colonnes_numeriques].std(skipna=True)
        print(f"   Vérification de la stabilité:")
        
        colonnes_importantes = ['T_AMBIANCE_01','T_AIR_E_FILTRE_A01','T_AIR_S_FILTRE_A02','T_AIR_S_TURBO_A03','T_AIR_E_MOTEUR_A04','T_FUEL_E_MOTEUR_A05','T_FUEL_E_RADIA_A06','T_FUEL_S_RADIA_A07','T_EAU_S_MOTEUR_A08', 
                                'T_EAU_E_MOTEUR_A09','EngineOilTemperature','T_HUILE_TRANS_A11','T_GAZ_ECHAPPEMENT_A15','R_CS.QFUKGH','R_EC.TORQUE','EngSpeed']
        colonnes_a_verifier = [c for c in colonnes_importantes if c in colonnes_numeriques]
        
        for col in colonnes_a_verifier:
            if col in ecarts_types.index and moyennes[col].iloc[0] != 0:
                cv = (ecarts_types[col] / abs(moyennes[col].iloc[0])) * 100
                
                if cv < 5:
                    print(f"      ✅ {col}: STABLE (CV={cv:.2f}%)")
                elif cv < 10:
                    print(f"      ⚠️ {col}: MOYENNEMENT STABLE (CV={cv:.2f}%)")
                else:
                    print(f"      ❌ {col}: INSTABLE (CV={cv:.2f}%)")
        
        moyennes['fichier_source'] = os.path.basename(fichier)
        regime = extraire_regime(os.path.basename(fichier))
        moyennes['regime_moteur'] = regime
        moyennes_fichiers.append(moyennes)
    
    if len(moyennes_fichiers) == 0:
        return None, None
    
    if 'fichier_source' not in colonnes_finales:
        colonnes_finales.append('fichier_source')
        unites_finales.append('')
    if 'regime_moteur' not in colonnes_finales:
        colonnes_finales.append('regime_moteur')
        unites_finales.append('tr/min')
    
    resultat_final = pd.concat(moyennes_fichiers, ignore_index=True, sort=False)
    resultat_final = resultat_final.reindex(columns=colonnes_finales)
    resultat_final = resultat_final.sort_values('regime_moteur').reset_index(drop=True)
    
    # Formules
    if all(col in resultat_final.columns for col in ['R_EC.TORQUE', 'K_TRA.RAPPORT_PDF']):
        resultat_final['Couple_moteur'] = resultat_final['R_EC.TORQUE'] / resultat_final['K_TRA.RAPPORT_PDF']
        colonnes_finales.append('Couple_moteur')
        unites_finales.append('N.m')

    if all(col in resultat_final.columns for col in ['T_AIR_E_MOTEUR_A04', 'K_TRA.T_AIR_MAXI']):
        resultat_final['TAA_AIR'] = resultat_final['K_TRA.T_AIR_MAXI'] - (resultat_final['T_AIR_E_MOTEUR_A04'] - resultat_final['T_AMBIANCE_01'])
        colonnes_finales.append('TAA_AIR')
        unites_finales.append('°C')

    if all(col in resultat_final.columns for col in ['EngineOilTemperature', 'K_TRA.T_OIL_MAXI']):
        resultat_final['TAA_HUILE'] = resultat_final['K_TRA.T_OIL_MAXI'] - (resultat_final['EngineOilTemperature'] - resultat_final['T_AMBIANCE_01'])
        colonnes_finales.append('TAA_HUILE')
        unites_finales.append('°C')

    if all(col in resultat_final.columns for col in ['T_EAU_S_MOTEUR_A08', 'K_TRA.T_EAU_MAXI']):
        resultat_final['TAA_EAU'] = resultat_final['K_TRA.T_EAU_MAXI'] - (resultat_final['T_EAU_S_MOTEUR_A08'] - resultat_final['T_AMBIANCE_01'])
        colonnes_finales.append('TAA_EAU')
        unites_finales.append('°C')

    if all(col in resultat_final.columns for col in ['EngSpeed', 'R_EC.TORQUE', 'K_TRA.RAPPORT_PDF']):
        resultat_final['Puissance_moteur'] = (resultat_final['EngSpeed'] * np.pi * 
                                              (resultat_final['R_EC.TORQUE'] / resultat_final['K_TRA.RAPPORT_PDF']) / 
                                              (30 * 1000))
        colonnes_finales.append('Puissance_moteur')
        unites_finales.append('kW')
    
    if 'Puissance_moteur' in resultat_final.columns and 'R_CS.QFUKGH' in resultat_final.columns:
        resultat_final['CSE_moteur'] = (resultat_final['R_CS.QFUKGH']*1000) / resultat_final['Puissance_moteur']
        colonnes_finales.append('CSE_moteur')
        unites_finales.append('g/kW.h')
    
    resultat_final = resultat_final.reindex(columns=colonnes_finales)
    
    print(f"\nAnalyse terminee: {len(resultat_final)} lignes")
    return resultat_final, (colonnes_finales, unites_finales)

def generer_dashboard_html(resultat_final, colonnes_info):
    colonnes_finales, unites_finales = colonnes_info
    colonnes_tableau = ['regime_moteur', 'AVG_PUISSANCE', 'T_AMBIANCE_01', 'R_CS.QFUKGH',
                        'TAA_AIR', 'TAA_EAU', 'TAA_HUILE','fichier_source']
    colonnes_tableau = [c for c in colonnes_tableau if c in resultat_final.columns]
    
    graphs_html = []
    
    categories_graphiques = {
        'Puissance': {
            'colonnes': ['AVG_PUISSANCE'],
            'axe_y': 'Puissance (kW)',
            'axe_y_min': None,
            'axe_y_max': None
        },
        'Couple': {
            'colonnes': ['R_EC.TORQUE', 'Couple_moteur'],
            'axe_y': 'Couple (N.m)',
            'axe_y_min': None,
            'axe_y_max': None
        },
        'Temperatures air': {
            'colonnes': ['T_AIR_E_FILTRE_A01', 'T_AIR_S_FILTRE_A02', 'T_AIR_S_TURBO_A03', 'T_AIR_E_MOTEUR_A04'],
            'axe_y': 'Temperature (°C)',
            'axe_y_min': 0,
            'axe_y_max': None
        },
        'Temperatures Fuel': {
            'colonnes': ['T_FUEL_E_MOTEUR_A05', 'T_FUEL_E_RADIA_A06', 'T_FUEL_S_RADIA_A07'],
            'axe_y': 'Temperature (°C)',
            'axe_y_min': 0,
            'axe_y_max': None
        },
        'Temperatures Eau/Huile': {
            'colonnes': ['T_EAU_S_MOTEUR_A08', 'T_EAU_E_MOTEUR_A09', 'T_FUEL_S_RADIA_A07', 'EngCoolanTemp','TCK_B01'],
            'axe_y': 'Temperature (°C)',
            'axe_y_min': 0,
            'axe_y_max': None
        },
        'Consommation': {
            'colonnes': ['C_CAL.CONSO','C_CAL.DEBIT_MASS','C_CAL.DEBIT_VOL', 'R_CS.QFUKGH'],
            'axe_y': 'Consommation',
            'axe_y_min': 0,
            'axe_y_max': None
        },
        'Pressions': {
            'colonnes': ['P_AIR_S_TURB','P_AIR_E_MOTEUR', 'P_EAU_S_MOTEUR', 'P_ECHAPPEMENT'],
            'axe_y': 'Pression (bar)',
            'axe_y_min': 0,
            'axe_y_max': None
        },
    }
    
    for categorie, config in categories_graphiques.items():
        cols_existantes = [c for c in config['colonnes'] if c in resultat_final.columns and resultat_final[c].notna().sum() > 0]
        if not cols_existantes:
            continue
        
        fig = go.Figure()
        for col in cols_existantes:
            idx_col = colonnes_finales.index(col) if col in colonnes_finales else -1
            unite = unites_finales[idx_col] if idx_col >= 0 and idx_col < len(unites_finales) else ''
            
            fig.add_trace(go.Scatter(
                x=resultat_final['regime_moteur'].tolist(),
                y=resultat_final[col].tolist(),
                mode='lines+markers',
                name=f"{col} ({unite})" if unite and unite != 'nan' else col,
                line=dict(width=2),
                marker=dict(size=8)
            ))
        
        yaxis_config = {
            'title': config['axe_y'],
            'gridcolor': '#e0e0e0',
            'zeroline': True,
            'zerolinecolor': '#888',
            'zerolinewidth': 1
        }
        
        if config['axe_y_min'] is not None:
            yaxis_config['range'] = [config['axe_y_min'], config['axe_y_max']]
        elif config['axe_y_max'] is not None:
            yaxis_config['range'] = [None, config['axe_y_max']]
        
        fig.update_layout(
            title=f"Evolution - {categorie}",
            xaxis_title="Regime moteur (tr/min)",
            yaxis=yaxis_config,
            template='plotly_white',
            height=500,
            hovermode='x unified',
            showlegend=True,
            legend=dict(
                orientation="v",
                yanchor="top",
                y=1,
                xanchor="left",
                x=1.02
            )
        )
        graphs_html.append(fig.to_json())
    
    html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Dashboard</title>'
    html += '<script src="https://cdn.plot.ly/plotly-2.26.0.min.js"></script>'
    html += '<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>'
    html += '<style>body{font-family:Arial;background:#667eea;padding:20px;}'
    html += '.container{max-width:1400px;margin:0 auto;background:white;border-radius:15px;padding:30px;}'
    html += 'h1{text-align:center;color:#2c3e50;}table{width:100%;border-collapse:collapse;}'
    html += 'th{background:#667eea;color:white;padding:10px;}td{padding:8px;border-bottom:1px solid #ddd;}'
    html += '.btn{background:#667eea;color:white;border:none;padding:10px 20px;border-radius:5px;cursor:pointer;}'
    html += '</style></head><body><div class="container">'
    html += f'<h1>Dashboard Analyse Moteur</h1><p style="text-align:center;color:#666;">Genere le {datetime.now().strftime("%d/%m/%Y à %H:%M")}</p>'
    
    html += '<button class="btn" onclick="exportToExcel()">Exporter en Excel</button>'
    html += '<table id="syntheseTable"><thead><tr>'
    for col in colonnes_tableau:
        html += f'<th>{col.replace("_", " ")}</th>'
    html += '</tr></thead><tbody>'
    
    for idx, row in resultat_final.iterrows():
        html += '<tr>'
        for col in colonnes_tableau:
            val = row[col]
            if pd.notna(val):
                if col == 'regime_moteur':
                    html += f'<td>{int(val)}</td>'
                elif col == 'fichier_source':
                    html += f'<td>{val}</td>'
                else:
                    html += f'<td>{val:.2f}</td>'
            else:
                html += '<td>-</td>'
        html += '</tr>'
    html += '</tbody></table>'
    
    for i, graph_json in enumerate(graphs_html):
        html += f'<div style="margin:30px 0;"><div id="graph{i}"></div></div>'
    
    html += '</div><script>function exportToExcel(){const table=document.getElementById("syntheseTable");'
    html += 'const wb=XLSX.utils.table_to_book(table);XLSX.writeFile(wb,"tableau_synthese.xlsx");}'
    for i, graph_json in enumerate(graphs_html):
        html += f'Plotly.newPlot("graph{i}",{graph_json}.data,{graph_json}.layout);'
    html += '</script></body></html>'
    
    return html

class InterfaceAnalyse:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Dashboard Analyse Moteur")
        self.window.geometry("900x700")
        self.fichiers_selectionnes = []
        self.creer_interface()
    
    def creer_interface(self):
        tk.Label(self.window, text="Dashboard Analyse Moteur", font=("Arial", 24, "bold")).pack(pady=20)
        tk.Label(self.window, text="Selectionnez vos fichiers Excel", font=("Arial", 11)).pack(pady=5)
        
        frame_boutons = tk.Frame(self.window)
        frame_boutons.pack(pady=30)
        
        tk.Button(frame_boutons, text="Selectionner un dossier", font=("Arial", 12), 
                  bg='#667eea', fg='white', padx=20, pady=15, 
                  command=self.selectionner_dossier).grid(row=0, column=0, padx=10)
        
        tk.Button(frame_boutons, text="Selectionner des fichiers", font=("Arial", 12),
                  bg='#764ba2', fg='white', padx=20, pady=15,
                  command=self.selectionner_fichiers).grid(row=0, column=1, padx=10)
        
        # Frame pour la période de moyennage
        frame_periode = tk.Frame(self.window)
        frame_periode.pack(pady=10)
        tk.Label(frame_periode, text="Moyennage sur les dernières:", 
                 font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
        self.periode_var = tk.IntVar(value=60)
        tk.Spinbox(frame_periode, from_=10, to=300, increment=10, 
                   textvariable=self.periode_var, width=5, 
                   font=("Arial", 10)).pack(side=tk.LEFT)
        tk.Label(frame_periode, text="secondes", 
                 font=("Arial", 9), fg='#666').pack(side=tk.LEFT, padx=5)
        
        frame_liste = tk.Frame(self.window, bg='white', relief=tk.SUNKEN, bd=2)
        frame_liste.pack(pady=20, padx=40, fill=tk.BOTH, expand=True)
        
        tk.Label(frame_liste, text="Fichiers selectionnes:", font=("Arial", 11, "bold"), 
                 bg='white').pack(pady=10, padx=10)
        
        scrollbar = tk.Scrollbar(frame_liste)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox = tk.Listbox(frame_liste, font=("Arial", 10), yscrollcommand=scrollbar.set)
        self.listbox.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)
        
        self.label_compteur = tk.Label(frame_liste, text="0 fichier(s)", font=("Arial", 10), bg='white')
        self.label_compteur.pack(pady=5)
        
        self.btn_analyser = tk.Button(self.window, text="Analyser les fichiers", 
                                       font=("Arial", 14, "bold"), bg='#28a745', fg='white',
                                       padx=40, pady=15, command=self.lancer_analyse, state=tk.DISABLED)
        self.btn_analyser.pack(pady=20)
    
    def selectionner_dossier(self):
        dossier = filedialog.askdirectory(title="Selectionner le dossier")
        if dossier:
            fichiers = glob.glob(os.path.join(dossier, "*.xlsx")) + glob.glob(os.path.join(dossier, "*.txt"))
            if len(fichiers) == 0:
                messagebox.showwarning("Aucun fichier", "Aucun fichier .xlsx trouve")
                return
            self.fichiers_selectionnes = fichiers
            self.afficher_fichiers()
    
    def selectionner_fichiers(self):
        fichiers = filedialog.askopenfilenames(title="Selectionner les fichiers Excel",
                                                filetypes=[("Fichiers Excel", "*.xlsx"), ("Fichiers Texte", "*.txt")])
        if fichiers:
            self.fichiers_selectionnes = list(fichiers)
            self.afficher_fichiers()
    
    def afficher_fichiers(self):
        self.listbox.delete(0, tk.END)
        for fichier in self.fichiers_selectionnes:
            self.listbox.insert(tk.END, os.path.basename(fichier))
        self.label_compteur.config(text=f"{len(self.fichiers_selectionnes)} fichier(s)")
        self.btn_analyser.config(state=tk.NORMAL if len(self.fichiers_selectionnes) > 0 else tk.DISABLED)
    
    def lancer_analyse(self):
        if len(self.fichiers_selectionnes) == 0:
            messagebox.showwarning("Aucun fichier", "Selectionnez des fichiers")
            return
        self.creer_fenetre_progression()
        thread = threading.Thread(target=self.executer_analyse)
        thread.start()
    
    def creer_fenetre_progression(self):
        self.fenetre_prog = tk.Toplevel(self.window)
        self.fenetre_prog.title("Analyse en cours...")
        self.fenetre_prog.geometry("400x150")
        self.fenetre_prog.transient(self.window)
        self.fenetre_prog.grab_set()
        
        tk.Label(self.fenetre_prog, text="Analyse en cours...", font=("Arial", 14, "bold")).pack(pady=20)
        self.label_progression = tk.Label(self.fenetre_prog, text="Traitement...", font=("Arial", 10))
        self.label_progression.pack(pady=10)
        self.progress_bar = ttk.Progressbar(self.fenetre_prog, length=300, mode='indeterminate')
        self.progress_bar.pack(pady=20)
        self.progress_bar.start(10)
    
    def executer_analyse(self):
        try:
            periode = self.periode_var.get()
            resultat_final, colonnes_info = analyser_fichiers_liste(
                self.fichiers_selectionnes, 
                periode_secondes=periode
            )
            if resultat_final is None:
                self.window.after(0, lambda: self.afficher_erreur("Echec de l'analyse"))
                return
            
            colonnes_finales, unites_finales = colonnes_info
            df_unites = pd.DataFrame([unites_finales], columns=colonnes_finales)
            df_export = pd.concat([pd.DataFrame([colonnes_finales], columns=colonnes_finales),
                                    df_unites, resultat_final], ignore_index=True)
            
            nom_fichier_excel = "fichier_concatene_moyennes_complet.xlsx"
            df_export.to_excel(nom_fichier_excel, index=False, header=False)
            
            html_content = generer_dashboard_html(resultat_final, colonnes_info)
            nom_fichier_html = "dashboard_analyse_moteur.html"
            with open(nom_fichier_html, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            self.window.after(0, lambda: self.afficher_succes(nom_fichier_excel, nom_fichier_html, 
                                                              len(resultat_final)))
        except Exception as e:
            self.window.after(0, lambda: self.afficher_erreur(str(e)))
    
    def afficher_succes(self, fichier_excel, fichier_html, nb_lignes):
        self.progress_bar.stop()
        self.fenetre_prog.destroy()
        message = f"Analyse terminee!\n\nFichiers:\n{fichier_excel}\n{fichier_html}\n\n"
        message += f"{nb_lignes} lignes\n\nOuvrir le dashboard?"
        if messagebox.askyesno("Termine", message):
            import webbrowser
            webbrowser.open(fichier_html)
    
    def afficher_erreur(self, erreur):
        self.progress_bar.stop()
        self.fenetre_prog.destroy()
        messagebox.showerror("Erreur", f"Erreur: {erreur}")
    
    def run(self):
        self.window.mainloop()

if __name__ == '__main__':
    print("Lancement de l'interface...")
    app = InterfaceAnalyse()
    app.run()