import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
from tkinter import ttk
import matplotlib.pyplot as plt


class Application:
    def __init__(self):# creer la fenetre principale et initialiser les variables
        #crée un dossier Sauvegardes/ si il n’existe pas
        self.dossier_sauvegarde = "Sauvegardes"
        os.makedirs(self.dossier_sauvegarde, exist_ok=True)

        # Données
        self.df = None

        # Historique des actions pour annuler/refaire

        """Le principe est le suivant :

Avant chaque modification du DataFrame, on sauvegarde une copie.
Si l'utilisateur clique sur Annuler, on revient à la copie précédente.
Si l'utilisateur clique sur Rétablir, on réapplique la modification annulée."""

        self.undo_stack = []
        self.redo_stack = []
        
        

            # Fenêtre principale
        self.fenetre = tk.Tk()
        self.fenetre.title("Gestionnaire de Données")
        # Centrer fenetre sur l'écran
        # taille générale de la fenêtre
        largeur_fenetre = 800 # largeur de la fenêtre
        hauteur_fenetre = 600 # hauteur de la fenêtre
        largeur_ecran = self.fenetre.winfo_screenwidth() # largeur de l'écran
        hauteur_ecran = self.fenetre.winfo_screenheight() # hauteur de l'écran
        position_x = (largeur_ecran - largeur_fenetre) // 2 # position x de la fenêtre
        position_y = (hauteur_ecran - hauteur_fenetre) // 2 # position y de la fenêtre
        self.fenetre.geometry(f"{largeur_fenetre}x{hauteur_fenetre}+{position_x}+{position_y}") # position de la fenêtre

        # couleur de fond de la fenêtre
        self.fenetre.configure(bg="lightblue") # couleur de fond de la fenêtre
        # transparence de la fenêtre
        self.fenetre.attributes("-alpha", 1) # transparence de la fenêtre

        # Treeview des fichiers sauvegardés
        self.tree_fichiers = ttk.Treeview(
            self.fenetre,
            columns=("Nom",),
            show="headings",
            height=20
        )
        # Treeview des données
        self.tree = ttk.Treeview(self.fenetre, show="headings") # tableau Excel où tu vois les données

        # barre de défilement verticale et orizontale
        ## verticale
        self.scrollbar_y=ttk.Scrollbar(self.fenetre, orient="vertical", command=self.tree.yview)
         ## horizontal
        self.scrollbar_x=ttk.Scrollbar(self.fenetre, orient="horizontal", command=self.tree.xview)


       

        # Création de l'interface
        self.creer_menu()
        self.creer_widgets()
        self.actualiser_fichiers()

    def creer_menu(self):
        # creer la barre de menu en haut comme excel. c’est juste la barre de commandes de ton logiciel
        menu_bar = tk.Menu(self.fenetre)

        # Menu Fichier
        menu_fichier = tk.Menu(menu_bar, tearoff=0)

        # Ajouter les commandes du menu Fichier
        menu_fichier.add_command(
            label="Charger",
            command=self.charger_donnees
        )
        menu_fichier.add_command(
            label="Nouveau_fenetre"       
        )
        menu_fichier.add_command(
            label="Exporter",
            command=self.exporter_affichage_dans_machine
        )
        menu_fichier.add_command(
            label="Sauvegarder",
            command=self.sauvegarder_donnees
        )
        # Ajouter un séparateur
        menu_fichier.add_separator()

        menu_fichier.add_command(
            label="Quitter",
            command=self.fenetre.quit
        )

        menu_bar.add_cascade(
            label="Fichier",
            menu=menu_fichier
        )

        # Menu Affichage
        menu_affichage = tk.Menu(menu_bar, tearoff=0)
        menu_affichage.add_command(
            label="Afficher les données",
            command=self.afficher_donnees
        )
        menu_bar.add_cascade(
            label="Affichage",
            menu=menu_affichage
        )

###menu edition 

        menu_edition = tk.Menu(menu_bar, tearoff=0)
        menu_atteindre = tk.Menu(menu_edition, tearoff=0)

        menu_atteindre.add_command(
            label="Recherche",
            command=self.rechercher
        )
        menu_atteindre.add_command(
            label="Recherche et remplacer",
            command=self.rechercher_et_remplacer
            )
        menu_edition.add_cascade(
            label="Atteindre",
            menu=menu_atteindre
        )

        menu_bar.add_cascade(
            label="Edition",
            menu=menu_edition
        )


        # Menu Traitement

        menu_traitement = tk.Menu(menu_bar, tearoff=0)

    # ---------- Sous-menu Colonnes ----------

        menu_colonnes = tk.Menu(menu_traitement, tearoff=0)
        menu_ajout_colonne = tk.Menu(menu_colonnes, tearoff=0)

        menu_ajout_colonne.add_command(
            label="Ajouter colonne simple",
            command=self.ajouter_colonne_simple
        )

        menu_ajout_colonne.add_command(
            label="Ajouter colonne avec formule",
            command=self.ajouter_colonne_avec_formule
        )

        menu_ajout_colonne.add_command(
            label="Ajouter colonne CRM",
            command=self.ajouter_colonne_CRM
        )

        menu_ajout_colonne.add_command(
            label="Ajouter colonne Blank",
            command=self.ajouter_colonne_blank
        )

        menu_colonnes.add_cascade(
            label="Ajouter",
            menu=menu_ajout_colonne
        )

        menu_colonnes.add_command(
            label="Supprimer une colonne",
            command=self.supprimer_colonnes
        )

        menu_colonnes.add_command(
            label="Sélectionner des colonnes",
            command=self.selectionner_colonnes
        )
        menu_colonnes.add_command(
            label="Renommer une colonne",
            command=self.renommer_colonne
        )

        menu_traitement.add_cascade(
            label="Colonnes",
            menu=menu_colonnes
        )

        # ---------- Sous-menu Lignes ----------

        menu_lignes = tk.Menu(menu_traitement, tearoff=0)

        menu_lignes.add_command(
            label="Filtrer les données",
            command=self.filtrer_donnees
        )

        menu_traitement.add_cascade(
            label="Lignes",
            menu=menu_lignes
        )

        # ---------- Sous-menu Fichiers ----------
  

        menu_fusion = tk.Menu(menu_traitement, tearoff=0)
        menu_jointure = tk.Menu(menu_fusion, tearoff=0)

        menu_jointure.add_command(
            label="jointure gauche",
            command=self.jointure_gauche)
        
        menu_jointure.add_command(
            label="jointure droite",
            command=self.jointure_droite)
        
        menu_jointure.add_command(
            label="jointure interne",
            command=self.jointure_interne)
        menu_jointure.add_command(
            label="jointure externe",
            command=self.jointure_externe
        )
        menu_fusion.add_cascade(
            label="Jointures",
            menu=menu_jointure)
        
        
        menu_fusion.add_separator()

        menu_fusion.add_command(
            label="Concaténer les fichiers", 
            command=self.concatener_fichiers
            )
        menu_traitement.add_cascade(
            label="Fusion des fichiers",
            menu=menu_fusion
            )
        
        menu_bar.add_cascade(
            label="Traitement",
            menu=menu_traitement
            )
        

### menu statistique
        menu_statistique = tk.Menu(menu_bar, tearoff=0)

        menu_statistique.add_command(
            label="type de donnée",
            command=self.type_donnée
        )

        menu_statistique.add_command(
            label="nombre occurrence objet",
            command=self.nombre_occurrence_objet
        )

        menu_statistique.add_command(
            label="statistique descriptive",
            command=self.statitique_descriptive
        )

        menu_bar.add_cascade(
            label="Statistique",
            menu=menu_statistique)

### menu graphique
        menu_graphique = tk.Menu(menu_bar, tearoff=0)
        menu_courbe=tk.Menu(menu_graphique, tearoff=0)

        menu_courbe.add_command(
            label="courbe from dataframe",
            command=self.courbe_from_dataframe
        )

        menu_courbe.add_command(
            label="courbe blank",
            command=self.courbe_blank
        )
        
        menu_graphique.add_cascade(
            label="Courbes",
            menu=menu_courbe
        )

        menu_graphique.add_command(
            label="histogramme"
        )

        menu_graphique.add_command(
            label="diagramme en barres"
        )

        menu_bar.add_cascade(
            label="Graphiques",
            menu=menu_graphique
        )

        self.fenetre.config(menu=menu_bar)
# -------------------------
    # WIDGETS
    # -------------------------

    def creer_widgets(self): # Créer les éléments visibles (widgets)
        

        # Titre
        titre = tk.Label(
            self.fenetre,
            text="Gestionnaire de Données Excel / CSV",
            font=("Arial", 16, "bold")
        )

        titre.grid(row=0, column=0, columnspan=2, pady=10)

        # ---------- Treeview des fichiers (épingle) ----------

        

        self.tree_fichiers.heading("Nom", text="Fichiers sauvegardés")
        self.tree_fichiers.column("Nom", width=220)  # liste des fichiers sauvegardés, 📌 colonne unique : Nom

        self.tree_fichiers.grid(
            row=1,
            column=0,
            padx=10,
            pady=10,
            sticky="ns"
        )
        self.tree_fichiers.bind("<Double-1>", self.ouvrir_fichier) #👉 double clic = ouvrir fichier

        # ---------- Treeview des données ----------
        self.tree.grid(
            row=1,
            column=1,
            padx=10,
            pady=10,
            sticky="nsew"
        ) #tableau Excel où tu vois les données

        # Permet au Treeview des données de s'agrandir. permet de redimensionner correctement
        self.fenetre.grid_rowconfigure(1, weight=1)
        self.fenetre.grid_columnconfigure(1, weight=1)

        self.bouton_supprimer = tk.Button(
            self.fenetre,
            text="Supprimer fichier",
            command=self.supprimer_fichier
        )
        # LIASON DES BARRE DE DEFILEMENT AVEC LE TREEVIEW DES DONNEES et placement
        self.tree.config(yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set)
        self.scrollbar_y.grid(row=1, column=2, sticky="ns")
        self.scrollbar_x.grid(row=2, column=1, sticky="ew")

        # BOUTON SUPPRIMER FICHIER
        self.bouton_supprimer.grid(
            row=2,
            column=0,
            padx=10,
            pady=10,
            sticky="w"
        )
        # BOUTON VIDER DOSSIER
        self.bouton_vider = tk.Button(
            self.fenetre,
            text="Vider dossier",
            command=self.vider_fichiers
        )
        self.bouton_vider.grid(
            row=2,
            column=0,
            padx=10,
            pady=10,
            sticky="e"
        )

        # BOUTON ANNULER ET RETABLIR
        ## bouton annuler
        self.bouton_annuler = tk.Button(
            self.fenetre,
            text="Annuler",
            command=self.annuler
        )
        self.bouton_annuler.grid(
            row=3,
            column=1,
            padx=10,
            pady=10,
            sticky="w"
        )

        ## bouton retablir
        self.bouton_retablir = tk.Button(
            self.fenetre,
            text="Rétablir",
            command=self.retablir
        )
        self.bouton_retablir.grid(
            row=3,
            column=1,
            padx=10,
            pady=10,
            sticky="e"
        )


    def actualiser_fichiers(self): # Afficher tous les fichiers dans le dossier Sauvegardes

        self.tree_fichiers.delete(*self.tree_fichiers.get_children()) #supprime anciens éléments

        for fichier in os.listdir(self.dossier_sauvegarde): # 👉 récupère tous les fichiers

            if fichier.endswith((".csv", ".xlsx")): # 👉 ajoute uniquement CSV et Excel

                self.tree_fichiers.insert(
                    "",
                    "end",
                    values=(fichier,)
                )
    def ouvrir_fichier(self, event): # ouvrir depuis la liste. Double clic → charger un fichier

        selection = self.tree_fichiers.selection()

        if not selection:
            return

        nom = self.tree_fichiers.item(selection[0])["values"][0]

        chemin = os.path.join(self.dossier_sauvegarde, nom)

        try:
            if nom.endswith(".csv"):
                self.df = pd.read_csv(chemin)
            else:
                self.df = pd.read_excel(chemin)

            self.afficher_donnees()

        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def charger_donnees(self): # ouvrir fichier externe

        fichier = filedialog.askopenfilename(
            title="Choisir un fichier",
            filetypes=[
                ("Fichiers Excel", "*.xlsx *.xls"),
                ("Fichiers CSV", "*.csv")
                ]
            )

        if not fichier:
            return

        try:

            if fichier.endswith((".xlsx", ".xls")):

                self.df = pd.read_excel(fichier)

            elif fichier.endswith(".csv"):
                self.df = pd.read_csv(
                        fichier,
                        sep=None, engine="python"
                        )

            """print(self.df.head())  # Affiche les premières lignes du DataFrame dans la console
            print(self.df.columns)  # Affiche les noms des colonnes dans la console"""

            self.afficher_donnees()

            messagebox.showinfo(
                "Succès",
                "Fichier chargé avec succès"
                )

        except Exception as e:

            messagebox.showerror(
                "Erreur",
                str(e)
                )
            
    def exporter_affichage_dans_machine(self):
        if self.df is None:
            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )
            return

        fichier = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel", "*.xlsx"),
                ("CSV", "*.csv")
            ]
        )

        if not fichier:
            return

        try:
            if fichier.endswith(".csv"):
                self.df.to_csv(fichier, index=False)
            else:
                self.df.to_excel(fichier, index=False)

            messagebox.showinfo(
                "Succès",
                "Fichier exporté avec succès"
            )

        except Exception as e:
            messagebox.showerror(
                "Erreur",
                str(e)
            )

     # -------------------------
    # SAUVEGARDE
    # -------------------------

    def sauvegarder_donnees(self): 

        #sauvegarde dans le dossier "Sauvegardes"
        if self.df is None:

            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )

            return
        nom = simpledialog.askstring(
            "Nom du fichier",
            "Nom du fichier :"
        )

        if not nom:
            return

        chemin = os.path.join(self.dossier_sauvegarde, nom + ".csv")

        self.df.to_csv(chemin, index=False)

        messagebox.showinfo("Succès", "Fichier sauvegardé")

        self.actualiser_fichiers()


# -------------------------
    # AFFICHAGE
    # -------------------------

    # Déclare une méthode qui affichera le contenu de self.df dans le Treeview. Cette méthode :

        """Vérifie que les données existent.
        Vide le tableau.
        Crée les colonnes.
        Affiche toutes les lignes du DataFrame dans le Treeview."""

    def afficher_donnees(self): # Affiche les données dans le tableau (Treeview) avec les colonnes et les lignes du DataFrame. Cette méthode est appelée après le chargement d'un fichier ou après toute modification des données. cet affichage est dynamique et s'adapte aux changements du DataFrame. Elle permet à l'utilisateur de visualiser les données de manière claire et organisée. la colonne index est ajoutée pour que l'utilisateur puisse voir le numéro de ligne.
        # Si aucun fichier n'a été chargé, la méthode s'arrête

        if self.df is None:
            return 
        
        # Effacer les anciennes données. Vider le tableau
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Définir les colonnes. Crée les colonnes du tableau à partir des colonnes du DataFrame et affiche leurs en-têtes.
        
        self.tree["columns"] = list(self.df.columns)
        self.tree["show"] = "tree headings"

        # Créer les en-têtes. Donne un nom à chaque colonne et fixe sa largeur

            # creer l'entete index
        self.tree.heading("#0", text="Index")
        self.tree.column("#0",width=50, stretch=False) 

            # creer les autre entete
        for col in self.df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col,width=80, stretch=False)

        # Ajouter les lignes. Parcourt les lignes du DataFrame et les ajoute dans le Treeview.
        for index, ligne in self.df.iterrows():
            self.tree.insert("", "end", text=index, values=list(ligne))


        
           
        






# -------------------------
    # EDITION
    # -------------------------
    """# recherche dans le tableau.
    def rechercher(self):
        if self.df is None:
            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )
            return

        recherche = simpledialog.askstring(
            "Recherche",
            "Texte à rechercher :"
        )

        if not recherche:
            return

        resultats = self.df.apply(lambda x: x.astype(str).str.contains(recherche, case=False)).any(axis=1)

        if resultats.any():
            self.df = self.df[resultats]
            self.afficher_donnees()
            messagebox.showinfo(
                "Succès",
                f"{resultats.sum()} résultats trouvés pour '{recherche}'"
            )
        else:
            messagebox.showinfo(
                "Résultat",
                f"Aucun résultat trouvé pour '{recherche}'"
            )"""

    # recherche dans une colonne specifique. L'utilisateur doit entrer le nom de la colonne et le texte à rechercher. Le programme doit vérifier que la colonne existe et afficher un message d'erreur si ce n'est pas le cas. Si la colonne existe le programme doit donner le nombre de résultat trouvé.
    def rechercher(self):
        if self.df is None:
            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )
            return

        colonne = simpledialog.askstring(
            "Recherche",
            f"Colonnes disponibles :\n\n{list(self.df.columns)}\n\nColonne à rechercher :"
        )

        if not colonne:
            return

        if colonne not in self.df.columns:
            messagebox.showerror(
                "Erreur",
                f"Colonne introuvable : {colonne}"
            )
            return

        recherche = simpledialog.askstring(
            "Recherche",
            "Texte à rechercher :"
        )

        if not recherche:
            return

        resultats = self.df[colonne].astype(str).str.contains(recherche, case=False)

        if resultats.any():
            self.df = self.df[resultats]
            self.afficher_donnees()
            messagebox.showinfo(
                "Succès",
                f"{resultats.sum()} résultats trouvés pour '{recherche}' dans la colonne '{colonne}'"
            )
        else:
            messagebox.showinfo(
                "Résultat",
                f"Aucun résultat trouvé pour '{recherche}' dans la colonne '{colonne}'"
            )
    
    # recherche et remplacer dans une colonne specifique. L'utilisateur doit entrer le nom de la colonne, le texte à rechercher et le texte de remplacement. Le programme doit vérifier que la colonne existe et afficher un message d'erreur si ce n'est pas le cas. Si la colonne existe le programme doit donner le nombre de résultat trouvé et remplacé.
    def rechercher_et_remplacer(self):
        if self.df is None:
            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )
            return

        colonne = simpledialog.askstring(
            "Recherche et remplacement",
            f"Colonnes disponibles :\n\n{list(self.df.columns)}\n\nColonne à rechercher :"
        )

        if not colonne:
            return

        if colonne not in self.df.columns:
            messagebox.showerror(
                "Erreur",
                f"Colonne introuvable : {colonne}"
            )
            return

        recherche = simpledialog.askstring(
            "Recherche et remplacement",
            "Texte à rechercher :"
        )

        if not recherche:
            return

        remplacement = simpledialog.askstring(
            "Recherche et remplacement",
            f"Texte de remplacement pour '{recherche}' :"
        )

        if remplacement is None:
            return

        resultats = self.df[colonne].astype(str).str.contains(recherche, case=False)
        nombre_resultats = resultats.sum()

        if nombre_resultats > 0:
            self.sauvegarder_etat()
            self.df.loc[resultats, colonne] = self.df.loc[resultats, colonne].astype(str).str.replace(recherche, remplacement, case=False)
            self.afficher_donnees()
            messagebox.showinfo(
                "Succès",
                f"{nombre_resultats} résultats trouvés et remplacés dans la colonne '{colonne}'"
            )
        else:
            messagebox.showinfo(
                "Résultat",
                f"Aucun résultat trouvé pour '{recherche}' dans la colonne '{colonne}'"
            )

# -------------------------
    # TRAITEMENT DES DONNÉES
    # -------------------------

    def ajouter_colonne_simple(self):
        if self.df is None:
            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )
            return

        nom_colonne = simpledialog.askstring(
            "Nouvelle colonne",
            "Nom de la nouvelle colonne :"
        )

        if not nom_colonne:
            return

        valeur_par_defaut = simpledialog.askstring(
            "Nouvelle colonne",
            f"Valeur par défaut pour la colonne '{nom_colonne}' :"
        )
        if valeur_par_defaut is None:
            return
        

        self.sauvegarder_etat()
        self.df[nom_colonne] = valeur_par_defaut
        self.afficher_donnees()

        messagebox.showinfo(
            "Succès",
            f"Colonne '{nom_colonne}' ajoutée avec la valeur par défaut '{valeur_par_defaut}'"
        )

    def ajouter_colonne_avec_formule(self):
        if self.df is None:
            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )
            return

        nom_colonne = simpledialog.askstring(
            "Nouvelle colonne avec formule",
            "Nom de la nouvelle colonne :"
        )

        if not nom_colonne:
            return

        formule = simpledialog.askstring(
            "Nouvelle colonne avec formule",
            f"Formule pour la colonne '{nom_colonne}' (ex: Col1 + Col2) :"
        )

        if not formule:
            return

        try:
            self.sauvegarder_etat()
            self.df[nom_colonne] = self.df.eval(formule)
            self.afficher_donnees()

            messagebox.showinfo(
                "Succès",
                f"Colonne '{nom_colonne}' ajoutée avec la formule '{formule}'"
            )
        except Exception as e:
            messagebox.showerror(
                "Erreur",
                f"Formule invalide : {str(e)}"
            )

    # ajouter colonne CRM. ajouter deux colonne nommé teneur_moy et ecart_typ et demander à l'utilisateur de donner les valeur par défaut pour ces deux colonne. ajouter quatre autres colonne nommée tm_plus_3et, tm_plus_2et, tm_moins_2et, tm_moins_3et et les valeur par défaut sont calculé automatiquement à partir des deux première colonne. les valeurs sont calculé comme suit : tm_plus_3et = t_moyenne + 3*ecart_typ, tm_plus_2et = t_moyenne + 2*ecart_typ, tm_moins_2et = t_moyenne - 2*ecart_typ, tm_moins_3et = t_moyenne - 3*ecart_typ. les colonnes sont nommées t_moyenne, ecart_typ, tm_plus_3et, tm_plus_2et, tm_moins_2et, tm_moins_3et

    #dans cette etépe on créee 6 colonnes.les nom des colonnes sont préétablit et sont t_moyenne, ecart_typ, tm_plus_3et, tm_plus_2et, tm_moins_2et, tm_moins_3et . l'utilisateur donne les deux premiere colonne qui sont t_moyenne(tm), ecart_typ(et) et les valeurs par défaut. les quatres autres colonnes sont calculées automatiquement t_moyenne +3*ecart_typ, t_moyenne +2*ecart_typ, t_moyenne -2*ecart_typ, t_moyenne -3*ecart_typ. les colonnes sont nommées t_moyenne, ecart_typ, tm_plus_3et, tm_plus_2et, tm_moins_2et, tm_moins_3et
    def ajouter_colonne_CRM(self):
        if self.df is None:
            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )
            return

        valeur_teneur_moy = simpledialog.askstring(
            "Nouvelle colonne CRM",
            "Valeur par défaut pour la colonne 'teneur_moy' :"
        )

        if valeur_teneur_moy is None:
            return

        valeur_ecart_typ = simpledialog.askstring(
            "Nouvelle colonne CRM",
            "Valeur par défaut pour la colonne 'ecart_typ' :"
        )

        if valeur_ecart_typ is None:
            return

        try:
            valeur_teneur_moy = float(valeur_teneur_moy)
            valeur_ecart_typ = float(valeur_ecart_typ)
        except ValueError:
            messagebox.showerror(
                "Erreur",
                "Les valeurs doivent être des nombres"
            )
            return

        self.sauvegarder_etat()
        self.df['teneur_moy'] = valeur_teneur_moy
        self.df['ecart_typ'] = valeur_ecart_typ
        self.df['tm_plus_3et'] = self.df['teneur_moy'] + 3 * self.df['ecart_typ']
        self.df['tm_plus_2et'] = self.df['teneur_moy'] + 2 * self.df['ecart_typ']
        self.df['tm_moins_2et'] = self.df['teneur_moy'] - 2 * self.df['ecart_typ']
        self.df['tm_moins_3et'] = self.df['teneur_moy'] - 3 * self.df['ecart_typ']
        self.afficher_donnees()

        messagebox.showinfo(
            "Succès",
            f"Colonnes 'teneur_moy' et 'ecart_typ' ajoutées avec les valeurs par défaut '{valeur_teneur_moy}' et '{valeur_ecart_typ}'"
        )


# pour blank : l'utilisateur ajoute au dataframe  deux colonnes prénommés. La premiére colonne s'appelle LD et la deuxiéme colonne s'appelle LDA. L'utilisateur doit entrer les valeurs par défaut pour ces deux colonnes
    def ajouter_colonne_blank(self):
        if self.df is None:
            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )
            return

        valeur_LD = simpledialog.askstring(
            "Nouvelle colonne Blank",
            "Valeur par défaut pour la colonne 'LD' (limite de détection) :"
        )

        if valeur_LD is None:
            return

        valeur_LDA = simpledialog.askstring(
            "Nouvelle colonne Blank",
            "Valeur par défaut pour la colonne 'LDA' (limite d'alerte) :"
        )

        
        if valeur_LDA is None:
            return
        

        try:
            valeur_LD = float(valeur_LD)
            valeur_LDA = float(valeur_LDA)
        except ValueError:
            messagebox.showerror(
                "Erreur",
                "Les valeurs doivent être des nombres"
            )
            return

        self.sauvegarder_etat()
        self.df['LD'] = valeur_LD
        self.df['LDA'] = valeur_LDA
        self.afficher_donnees()

        messagebox.showinfo(
            "Succès",
            f"Colonnes 'LD' et 'LDA' ajoutées avec les valeurs par défaut '{valeur_LD}' et '{valeur_LDA}'"
        )




    def supprimer_colonnes(self):
        if self.df is None:

            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )

            return

        colonnes = simpledialog.askstring(
            "Suppression",
            f"Colonnes disponibles :\n\n{list(self.df.columns)}\n\nColonnes à supprimer (séparées par des virgules) :"
        )

        if not colonnes:
            return

        colonnes = [col.strip() for col in colonnes.split(",")]

        for col in colonnes:
            if col not in self.df.columns:
                messagebox.showerror(
                    "Erreur",
                    f"Colonne introuvable : {col}"
                )
                return
            # self.sauvegarder_etat() Pour sauvegarder l'état avant la suppression des colonnes. anulable pour revenir en arrière si nécessaire.
        self.sauvegarder_etat()
        self.df.drop(columns=colonnes, inplace=True)

        self.afficher_donnees()

        messagebox.showinfo(
            "Succès",
            f"Colonnes supprimées : {colonnes}"
        )
    def selectionner_colonnes(self):
        if self.df is None:

            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )

            return

        colonnes = simpledialog.askstring(
            "Filtrage",
            f"Colonnes disponibles :\n\n{list(self.df.columns)}\n\nColonnes à conserver (séparées par des virgules) :"
        )

        if not colonnes:
            return

        colonnes = [col.strip() for col in colonnes.split(",")]

        for col in colonnes:
            if col not in self.df.columns:
                messagebox.showerror(
                    "Erreur",
                    f"Colonne introuvable : {col}"
                )
                return
        self.sauvegarder_etat()
        self.df = self.df[colonnes]

        self.afficher_donnees()

        messagebox.showinfo(
            "Succès",
            f"Colonnes conservées : {colonnes}"
        )

    def renommer_colonne(self):
        if self.df is None:

            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )

            return

        colonnes = simpledialog.askstring(
            "Renommer une colonne",
            f"Colonnes disponibles :\n\n{list(self.df.columns)}\n\nColonne à renommer :"
        )

        if not colonnes:
            return

        if colonnes not in self.df.columns:
            messagebox.showerror(
                "Erreur",
                f"Colonne introuvable : {colonnes}"
            )
            return

        nouveau_nom = simpledialog.askstring(
            "Renommer une colonne",
            f"Nouveau nom pour la colonne '{colonnes}' :"
        )

        if not nouveau_nom:
            return

        self.sauvegarder_etat()
        self.df.rename(columns={colonnes: nouveau_nom}, inplace=True)

        self.afficher_donnees()

        messagebox.showinfo(
            "Succès",
            f"Colonne '{colonnes}' renommée en '{nouveau_nom}'"
        )
    # filtrer les données selon une condition. Par exemple, si l'utilisateur veut filtrer les données pour ne garder que les lignes où la valeur de la colonne "Age" est supérieure à 30, il pourra entrer "Age > 30" dans une boîte de dialogue. Le programme appliquera ensuite ce filtre et affichera les résultats dans le tableau.
    def filtrer_donnees(self):
        if self.df is None:

            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )

            return

        condition = simpledialog.askstring(
            "Filtrage",
            f"Colonnes disponibles :\n\n{list(self.df.columns)}\n\nCondition de filtrage (ex: Age > 30) :"
        )

        if not condition:
            return

        try:
            self.sauvegarder_etat()
            self.df = self.df.query(condition)
            self.afficher_donnees()
            messagebox.showinfo(
                "Succès",
                f"Données filtrées selon la condition : {condition}"
            )
        except Exception as e:
            messagebox.showerror(
                "Erreur",
                f"Condition invalide : {str(e)}"
            )

    

    def fusionner_fichiers(self):
            pass

    
###### JOINTURE DEUX FICHIERS
    # chargement du deuxieme fuchier pour la fusion. Le programme doit demander à l'utilisateur de sélectionner un deuxième fichier (Excel ou CSV) et de choisir le type de fusion (concaténation ou jointure). Ensuite, il doit effectuer la fusion et afficher le résultat dans le tableau.
    def charger_df_secondaire(self):

        fichier = filedialog.askopenfilename(
            title="Choisir le deuxième fichier",
            filetypes=[
                ("Excel", "*.xlsx *.xls"),
                ("CSV", "*.csv")
            ]
        )

        if not fichier:
            return None

        try:
            if fichier.endswith(".csv"):
                return pd.read_csv(fichier)
            else:
                return pd.read_excel(fichier)

        except Exception as e:
            messagebox.showerror("Erreur", str(e))
            return None
        
    def jointure(self, type_join):

        if self.df is None:
            messagebox.showwarning("Attention", "Aucune donnée chargée")
            return

        df2 = self.charger_df_secondaire()
        if df2 is None:
            return

        # Colonnes (différentes autorisées)
        col1 = simpledialog.askstring(
            "Jointure",
            f"Colonnes fichier principal : {list(self.df.columns)}\n\nColonne fichier 1 :"
        )

        if not col1:
            return

        col2 = simpledialog.askstring(
            "Jointure",
            f"Colonnes fichier secondaire : {list(df2.columns)}\n\nColonne fichier 2 :"
        )

        if not col2:
            return

        try:
            self.sauvegarder_etat()
            self.df = pd.merge(
                self.df,
                df2,
                how=type_join,
                left_on=col1,
                right_on=col2
            )

            self.afficher_donnees()

            messagebox.showinfo("Succès", f"Jointure {type_join} effectuée")

        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def jointure_gauche(self):
        self.jointure("left")

    def jointure_droite(self):
        self.jointure("right")

    def jointure_interne(self):
        self.jointure("inner")

    def jointure_externe(self):
        self.jointure("outer")

    # concaténer les fichiers. Le programme doit demander à l'utilisateur de sélectionner un deuxième fichier (Excel ou CSV) et de choisir si la concaténation doit se faire par lignes ou par colonnes. Ensuite, il doit effectuer la concaténation et afficher le résultat dans le tableau.
    def concatener_fichiers(self):
        if self.df is None:
            messagebox.showwarning("Attention", "Aucune donnée chargée")
            return

        df2 = self.charger_df_secondaire()
        if df2 is None:
            return

        orientation = simpledialog.askstring(
            "Concaténation",
            "Concaténer par lignes (vertical) ou par colonnes (horizontal) ? (v/h)"
        )

        if orientation not in ["v", "h"]:
            messagebox.showerror("Erreur", "Orientation invalide")
            return

        try:
            self.sauvegarder_etat()
            if orientation == "v":
                self.df = pd.concat([self.df, df2], axis=0, ignore_index=True)
            else:
                self.df = pd.concat([self.df, df2], axis=1)

            self.afficher_donnees()

            messagebox.showinfo("Succès", "Concaténation effectuée")

        except Exception as e:
            messagebox.showerror("Erreur", str(e))

## statistique 
## d'abitude les fonction statistque envoi du texte ou des tableau donc le mieu c'est de creer une fenetre(Toplevel) qui vas afficher ces texte ou tableau dans une nouvelle fentre sans rapport avec notre fenetre peincipal.
    def fenetre_affichage_stat(self, titre, contenu):
        top_fenetre = tk.Toplevel(self.fenetre)
        top_fenetre.title(titre)
        top_fenetre.geometry("600x400")

        # Permet au widget Text de s'agrandir avec la fenêtre
        top_fenetre.rowconfigure(0, weight=1)
        top_fenetre.columnconfigure(0, weight=1)

        # Zone de texte
        texte_widget = tk.Text(top_fenetre, wrap="none")

        # Barre de défilement verticale
        scrollbar_toplevel_y = tk.Scrollbar(top_fenetre, orient="vertical",
                                command=texte_widget.yview)

        # Barre de défilement horizontal

        scrollbar_toplevel_x = tk.Scrollbar(top_fenetre, orient="vertical",
                                command=texte_widget.xview)
        
        texte_widget.configure(
        yscrollcommand=scrollbar_toplevel_y.set,
        xscrollcommand=scrollbar_toplevel_x.set
    )
        

        # Placement
        texte_widget.grid(row=0, column=0, sticky="nsew")
        scrollbar_toplevel_y.grid(row=0, column=1, sticky="ns")
        scrollbar_toplevel_x.grid(row=1, column=0, sticky="ew")


        # Insertion du contenu
        texte_widget.insert(tk.END, contenu)

        # Lecture seule
        texte_widget.config(state="disabled")
# fonction dtype
# creer une fonction qui génére le nombre de type de donnée présent dans notre donnée c'est a dire pour colonne X on a tel nombre de données int, float, str. 

    def type_donnée (self) :
        
        if self.df is None:
            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )
            return
               # Normalement, df.info() affiche son résultat directement dans la console. Grâce à StringIO, on peut récupérer ce texte dans une variable au lieu de l'envoyer à l'écran. 
        from io import StringIO
        buffer_info=StringIO()  # On crée un espace vide qui va recevoir le texte produit par df.info().On peut l'imaginer comme un cahier vide.
        self.df.info(buf=buffer_info) # Écrire les informations dans ce buffer c'est-à-dire dans notre cahier vide. Au lieu d'afficher le texte à l'écran, df.info() l'écrit dans buffer_info.


        self.fenetre_affichage_stat("Information sur type de donnée", buffer_info.getvalue()) # Afficher le résultat dans une fenêtre

    def statitique_descriptive(self):

        if self.df is None:
            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )
            return

        from io import StringIO
        buffer_stat_descriptive=StringIO()
        self.df.describe(include="all").to_string(buf=buffer_stat_descriptive)

        self.fenetre_affichage_stat("Information sur type de donnée", buffer_stat_descriptive.getvalue())


    # nombre d'occurence des colonne de type object. creer une fonction qui genere le nombre d'occurence de chaque valeur pour chaque colonne de type object. afficher le resultat dans une nouvelle sous forme de tableau avec la colonne de gauche le nom de la colonne et la colonne de droite le nombre d'occurence de chaque valeur. afficher le resultat dans une nouvelle fenetre avec un widget Text et une barre de defilement.
    def nombre_occurrence_objet(self):
        if self.df is None:
            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )
            return

        colonnes_objet = self.df.select_dtypes(include=["object"]).columns

        if len(colonnes_objet) == 0:
            messagebox.showinfo(
                "Information",
                "Aucune colonne de type 'object' trouvée"
            )
            return

        from io import StringIO
        buffer_occurrence=StringIO()

        for col in colonnes_objet:
            buffer_occurrence.write(f"Colonne : {col}\n")
            buffer_occurrence.write(self.df[col].value_counts().to_string())
            buffer_occurrence.write("\n\n")

        self.fenetre_affichage_stat("Nombre d'occurrences des colonnes de type 'object'", buffer_occurrence.getvalue())


        
# -------------------------
    # GRAPHIQUES
    # -------------------------
    # courbe avec matplotlib. plt.figure() pour creer plusieurs courbe dans la meme graphique. demander l'utilisateur de choisir une colonne pour l'axe des x et une ou plusieurs colonnes pour l'axe des y et preciser une des  colonne de y qu'on va representer comme dispersion.  afficher la courbe dans une nouvelle fenetre avec plt.show().
    def courbe_from_dataframe(self):
        if self.df is None:
            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )
            return

        colonne_x = simpledialog.askstring(
            "Graphique",
            f"Colonnes disponibles :\n\n{list(self.df.columns)}\n\nColonne pour l'axe des X :"
        )

        if not colonne_x:
            return

        colonnes_y = simpledialog.askstring(
            "Graphique",
            f"Colonnes disponibles :\n\n{list(self.df.columns)}\n\nColonnes pour l'axe des Y (séparées par des virgules) :"
        )

        if not colonnes_y:
            return

        colonnes_y = [col.strip() for col in colonnes_y.split(",")]

        for col in colonnes_y:
            if col not in self.df.columns:
                messagebox.showerror(
                    "Erreur",
                    f"Colonne introuvable : {col}"
                )
                return

        dispersion_colonne = simpledialog.askstring(
            "Graphique",
            f"Colonnes disponibles :\n\n{list(self.df.columns)}\n\nColonne pour la dispersion (optionnelle, séparée par des virgules, laisser vide si aucune) :"
        )

        if dispersion_colonne:
            dispersion_colonne = [col.strip() for col in dispersion_colonne.split(",")]
            for col in dispersion_colonne:
                if col not in self.df.columns:
                    messagebox.showerror(
                        "Erreur",
                        f"Colonne introuvable : {col}"
                    )
                    return

        plt.figure()

        for col in colonnes_y:
            plt.plot(self.df[colonne_x], self.df[col], label=col)

        if dispersion_colonne and dispersion_colonne in self.df.columns:
            plt.scatter(self.df[colonne_x], self.df[dispersion_colonne], color='red', label=dispersion_colonne)

        plt.xlabel(colonne_x)
        plt.ylabel("Valeurs")
        plt.title("Graphique en courbe")
        plt.legend()
        plt.grid()
        plt.show()

    def courbe_blank(self):
        if self.df is None:
            messagebox.showwarning(
                "Attention",
                "Aucune donnée chargée"
            )
            return

        colonne_x = simpledialog.askstring(
            "Graphique",
            f"Colonnes disponibles :\n\n{list(self.df.columns)}\n\nColonne pour l'axe des X :"
        )

        if not colonne_x:
            return
        if colonne_x not in self.df.columns:
            messagebox.showerror(
                "Erreur",
                f"Colonne introuvable : {colonne_x}"
            )
            return

        colonnes_y = simpledialog.askstring(
            "Graphique",
            f"Colonnes disponibles :\n\n{list(self.df.columns)}\n\nEntrer les colonnes correspondante à la limite de détection et la limite d'alerte (séparées par des virgules) :"
        )

        if not colonnes_y:
            return

        colonnes_y = [col.strip() for col in colonnes_y.split(",")]

        for col in colonnes_y:
            if col not in self.df.columns:
                messagebox.showerror(
                    "Erreur",
                    f"Colonne introuvable : {col}"
                )
                return

        dispersion_colonne = simpledialog.askstring(
            "Graphique",
            f"Colonnes disponibles :\n\n{list(self.df.columns)}\n\nEntrer la colonne correspondante à la teneur du blank :"
        )

        if not dispersion_colonne:
            
            return
        if dispersion_colonne not in self.df.columns:
            messagebox.showerror(
                "Erreur",
                f"Colonne introuvable : {dispersion_colonne}"
            )
            return

        plt.figure()

        for col in colonnes_y:
            plt.plot(self.df[colonne_x], self.df[col], label=col)

        if dispersion_colonne and dispersion_colonne in self.df.columns:
            plt.scatter(self.df[colonne_x], self.df[dispersion_colonne], color='red', label=dispersion_colonne)

        plt.xlabel(colonne_x)
        plt.ylabel("Valeurs")
        plt.title("Graphique blank")
        plt.legend()
        plt.grid()
        plt.show()

    





    # creer une fonction de sauvegarde pour sauvegarder chaque fichier avant chaque modification. Cette fonction doit être appelée avant chaque modification du DataFrame (ajout, suppression, modification de colonnes ou de lignes). Elle doit sauvegarder une copie du DataFrame actuel dans une pile (stack) pour permettre l'annulation des modifications. Si l'utilisateur effectue une nouvelle action après avoir annulé une modification, la pile de redo doit être effacée.

    def sauvegarder_etat(self):
        if self.df is not None:
            self.undo_stack.append(self.df.copy(deep=True))

            # une nouvelle action efface le redo
            self.redo_stack.clear()

                # fonction annuler pour revenir à l'état précédent du DataFrame. Cette fonction doit vérifier si la pile d'annulation (undo_stack) n'est pas vide. Si elle contient des états précédents, elle doit restaurer le DataFrame à l'état précédent et mettre à jour l'affichage. Si la pile est vide, un message d'information doit être affiché pour indiquer qu'aucune action ne peut être annulée.
    def annuler(self):

        if not self.undo_stack:
            messagebox.showinfo(
                "Information",
                "Aucune action à annuler."
            )
            return

        self.redo_stack.append(self.df.copy(deep=True))

        self.df = self.undo_stack.pop()

        self.afficher_donnees()
    
                # fonction retablir pour revenir à l'état suivant du DataFrame. Cette fonction doit vérifier si la pile de rétablissement (redo_stack) n'est pas vide. Si elle contient des états suivants, elle doit restaurer le DataFrame à l'état suivant et mettre à jour l'affichage. Si la pile est vide, un message d'information doit être affiché pour indiquer qu'aucune action ne peut être refaite.

    def retablir(self):

        if not self.redo_stack:
            messagebox.showinfo(
                "Information",
                "Aucune action à rétablir."
            )
            return

        self.undo_stack.append(self.df.copy(deep=True))

        self.df = self.redo_stack.pop()

        self.afficher_donnees()

   



    # creer creer un bouton supprimer pour supprimer un ou plusieurs fichiers dans le dossier Sauvegardes. Le bouton doit être placé sous la liste des fichiers sauvegardés. Quand l'utilisateur clique sur le bouton, le programme doit demander confirmation avant de supprimer le fichier sélectionné. Si aucun fichier n'est sélectionné, un message d'avertissement doit s'afficher.

    def supprimer_fichier(self):

        selection = self.tree_fichiers.selection()

        if not selection:
            messagebox.showwarning(
                "Attention",
                "Aucun fichier sélectionné"
            )
            return

        nom = self.tree_fichiers.item(selection[0])["values"][0]

        confirmation = messagebox.askyesno(
            "Confirmation",
            f"Voulez-vous vraiment supprimer le fichier '{nom}' ?"
        )

        if confirmation:
            chemin = os.path.join(self.dossier_sauvegarde, nom)
            os.remove(chemin)
            self.actualiser_fichiers()
            messagebox.showinfo(
                "Succès",
                f"Fichier '{nom}' supprimé"
            )

    # creer un bouton vider pour vider les fichier dans le dossier Sauvegardes. Le bouton doit être placé sous la liste des fichiers sauvegardés. Quand l'utilisateur clique sur le bouton, le programme doit demander confirmation avant de supprimer tous les fichiers dans le dossier Sauvegardes. Si aucun fichier n'est présent, un message d'avertissement doit s'afficher.

    def vider_fichiers(self):

        fichiers = os.listdir(self.dossier_sauvegarde)

        if not fichiers:
            messagebox.showwarning(
                "Attention",
                "Aucun fichier à supprimer"
            )
            return

        confirmation = messagebox.askyesno(
            "Confirmation",
            "Voulez-vous vraiment supprimer tous les fichiers ?"
        )

        if confirmation:
            for fichier in fichiers:
                chemin = os.path.join(self.dossier_sauvegarde, fichier)
                os.remove(chemin)
            self.actualiser_fichiers()
            messagebox.showinfo(
                "Succès",
                "Tous les fichiers ont été supprimés"
            )
        





    # -------------------------
    # EXECUTION
    # -------------------------


    def executer(self):

        self.fenetre.mainloop()

# Lancement du programme
if __name__ == "__main__":

    app = Application()
    app.executer()






