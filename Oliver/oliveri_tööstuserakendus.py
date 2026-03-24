#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import pandas as pd
import re


# Esimesed 6 veergu on toote pohilandmed
TOOTE_VEERUD = [
    "A/U",
    "toote tyyp",
    "laius",
    "korgus",
    "tk tyyptellimusel",
    "kasi",
]


class Tootmisrakendus:
    def __init__(self, fail="Tootmisreeglid.xlsx"):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        fail_path = os.path.join(script_dir, fail)

        # Loeme kahepaiselise tabelina (1. ja 2. rida on paised)
        self.df = pd.read_excel(fail_path, header=[0, 1], engine='openpyxl')

        # Puhastame paised
        self.df.columns = pd.MultiIndex.from_tuples(
            [
                (
                    str(l0).strip() if pd.notna(l0) else "",
                    str(l1).strip() if pd.notna(l1) else "",
                )
                for (l0, l1) in self.df.columns.to_list()
            ]
        )

        self.tooted = {}
        self._lae_tooted()

    # -----------------------------
    #  VALEMITE ARVUTAMINE
    # -----------------------------
    def _eval_valem(self, expr, L, K, tk, adjustments=None):
        """Arvutab valemi stringist, kasutades ainult L, K ja tk, ning rakendab dunaamilised kohandused."""
        if not isinstance(expr, str):
            return expr

        # Rakenda kohandused
        if adjustments:
            L = adjustments.get("laius", L)
            K = adjustments.get("korgus", K)
            tk = adjustments.get("tk", tk)

        s = expr.replace("L", str(L)).replace("K", str(K))
        s = s.replace("tkx4", str(tk * 4))
        s = s.replace("tkx2", str(tk * 2))
        s = s.replace("tk", str(tk))

        try:
            return eval(s)
        except:
            return expr

    def _parse_soovid(self, soovid):
        """Parsib kliendi soove ja tagastab kohanduste sonastiku."""
        adjustments = {}
        for soov in soovid:
            soov = soov.lower()
            
            # Eemalda klaasiliistud
            if "ilma klaasiliistudeta" in soov:
                adjustments["klaasiliistud"] = 0
            
            # Kohandatud klaasiliistud
            match = re.search(r"(\d+)\s+vertikaalset\s+ja\s+(\d+)\s+horisontaalset", soov)
            if match:
                adjustments["vertikaalsed_klaasiliistud"] = int(match.group(1))
                adjustments["horisontaalsed_klaasiliistud"] = int(match.group(2))
            
            # Kohandatud moot (naiteks "10 mm kitsam leng")
            match = re.search(r"(\d+)\s*mm\s*(kitsam|laiem)\s*leng", soov)
            if match:
                mm = int(match.group(1))
                if match.group(2) == "kitsam":
                    adjustments["leng_kohandus"] = -mm
                else:
                    adjustments["leng_kohandus"] = mm
        
        return adjustments

    # -----------------------------
    #  TOODETE LAADIMINE
    # -----------------------------
    def _lae_tooted(self):
        rows = self.df

        # Kaime ridade kaupa: numbrid (0), valemid (1), numbrid (2), valemid (3) jne
        for idx in range(0, len(rows), 2):
            num_row = rows.iloc[idx]

            # Kui valemirida puudub → lopetame
            if idx + 1 >= len(rows):
                break

            valem_row = rows.iloc[idx + 1]

            # Toote pohilandmed tulevad numbrilisest reast
            a_u = num_row.get(("A/U", "Unnamed: 0_level_1"), None)
            toote_tyyp = num_row.get(("toote tyyp", "Unnamed: 1_level_1"), None)

            if pd.isna(a_u) or pd.isna(toote_tyyp):
                continue

            # Need kolm vaartust on valemite alus
            L = float(num_row.get(("laius", "Unnamed: 2_level_1"), 0))
            K = float(num_row.get(("korgus", "Unnamed: 3_level_1"), 0))
            tk = int(num_row.get(("tk tyyptellimusel", "Unnamed: 4_level_1"), 0))

            kasi = num_row.get(("kasi", "Unnamed: 5_level_1"), None)

            detailid = {}

            # -----------------------------
            #  DETAILIDE DUNAAMILINE PARSIMINE
            # -----------------------------
            columns = list(self.df.columns)
            i = 0

            while i < len(columns):
                grp, field = columns[i]

                # Jata toote pohilandmed ja tyhjad grupid vahele
                if grp in TOOTE_VEERUD or grp == "":
                    i += 1
                    continue

                # Jata "konfig." taielikult valja
                if field.lower().startswith("konfig"):
                    i += 1
                    continue

                # Loo grupp kui seda pole
                if grp not in detailid:
                    detailid[grp] = {"mootud": {}, "kogus": {}}

                # Kui see on moodu veerg (mitte tk)
                if not field.lower().startswith("tk"):
                    valem = valem_row[(grp, field)]
                    if not pd.isna(valem):
                        detailid[grp]["mootud"][field] = self._eval_valem(valem, L, K, tk)

                    # Kontrolli, kas jargmine veerg on kogus
                    if i + 1 < len(columns):
                        next_grp, next_field = columns[i + 1]

                        if next_grp == grp and next_field.lower().startswith("tk"):
                            valem_kogus = valem_row[(next_grp, next_field)]
                            if not pd.isna(valem_kogus):
                                detailid[grp]["kogus"][field] = self._eval_valem(
                                    valem_kogus, L, K, tk
                                )

                            i += 2
                            continue

                    i += 1
                    continue

                # Kui on koguse veerg, aga moot oli juba toodeldud → edasi
                i += 1

            # Salvestame toote
            toote_entry = {
                "toote_tyyp": toote_tyyp,
                "A/U": a_u,
                "mootud": {"laius": L, "korgus": K, "tk": tk},
                "detailid": {"kasi": kasi, "grupid": detailid},
            }

            key = str(toote_tyyp)
            self.tooted.setdefault(key, []).append(toote_entry)

    # -----------------------------
    #  TOOTE DETAILIDE PARING
    # -----------------------------
    def get_toote_detailid(self, toote_tyyp, laius=None, soovid=None):

        if toote_tyyp not in self.tooted:
            raise ValueError("Toodet " + toote_tyyp + " ei leitud.")

        kandidaadid = self.tooted[toote_tyyp]

        if laius is None:
            toode = kandidaadid[0]
        else:
            for t in kandidaadid:
                if t["mootud"]["laius"] == float(laius):
                    toode = t
                    break
            else:
                raise ValueError("Toodet " + toote_tyyp + " laiusega " + str(laius) + " ei leitud.")

        # Kohanda toodet vastavalt kliendi soovidele
        if soovid:
            adjustments = self._parse_soovid(soovid)
            toode = self._apply_adjustments(toode, adjustments)

        return toode

    def _apply_adjustments(self, toode, adjustments):
        """Rakendab kohandused tootele."""
        toode = toode.copy()
        
        # Kohanda mootmed
        if "leng_kohandus" in adjustments:
            toode["mootud"]["laius"] += adjustments["leng_kohandus"]
            # Kohanda klaasi moot, kui leng on kitsam
            if "klaas" in toode["detailid"]["grupid"]:
                for key in toode["detailid"]["grupid"]["klaas"]["mootud"]:
                    if isinstance(toode["detailid"]["grupid"]["klaas"]["mootud"][key], (int, float)):
                        toode["detailid"]["grupid"]["klaas"]["mootud"][key] += adjustments["leng_kohandus"]
        
        # Kohanda klaasiliistud
        if "klaasiliistud" in adjustments:
            if "klaasiliistud" in toode["detailid"]["grupid"]:
                toode["detailid"]["grupid"]["klaasiliistud"]["kogus"] = adjustments["klaasiliistud"]
        
        if "vertikaalsed_klaasiliistud" in adjustments:
            if "klaasiliistud" in toode["detailid"]["grupid"]:
                toode["detailid"]["grupid"]["klaasiliistud"]["vertikaalsed"] = adjustments["vertikaalsed_klaasiliistud"]
        
        if "horisontaalsed_klaasiliistud" in adjustments:
            if "klaasiliistud" in toode["detailid"]["grupid"]:
                toode["detailid"]["grupid"]["klaasiliistud"]["horisontaalsed"] = adjustments["horisontaalsed_klaasiliistud"]
        
        return toode

    def get_all_tooted(self):
        return list(self.tooted.keys())


# -----------------------------
#  JSON VALJUND
# -----------------------------
if __name__ == "__main__":
    import json

    rakendus = Tootmisrakendus()

    # Kuva saadaolevad tooted
    print("Saadaolevad tooted:", rakendus.get_all_tooted())
    
    # Kui on tooted saadaolevad, proovi kohandada esimest toodet
    if rakendus.get_all_tooted():
        esimesed_toode_tyyp = rakendus.get_all_tooted()[0]
        print("Proovime kohandada toodet:", esimesed_toode_tyyp)
        
        # Leiame esimese toote laius
        esimesed_toode = rakendus.tooted[esimesed_toode_tyyp][0]
        esimesed_laius = esimesed_toode["mootud"]["laius"]
        
        toode = rakendus.get_toote_detailid(
            esimesed_toode_tyyp,
            laius=esimesed_laius,
            soovid=[
                "sooviks akent ilma klaasiliistudeta",
                "soovin 2 vertikaalset ja uhe horisontaalse klaasiliistu",
                "soovin, et leng oleks 10 mm kitsam",
            ],
        )

        print("Kohandatud toode:")
        print(json.dumps(toode, indent=2, ensure_ascii=False))
    else:
        print("Tooteid ei leitud!")

    # Salvestame koik tooted
    tooted_json = {}
    for toode_nimi in rakendus.get_all_tooted():
        for variant in rakendus.tooted[toode_nimi]:
            key = '{0}_{1}x{2}'.format(toode_nimi, int(variant["mootud"]["laius"]), int(variant["mootud"]["korgus"]))
            tooted_json[key] = variant

    with open("tootmisreeglid_output.json", "w", encoding="utf-8") as f:
        json.dump(tooted_json, f, indent=2, ensure_ascii=False)

    print("\nTootmisreeglid on salvestatud faili 'tootmisreeglid_output.json'.")
