#!/usr/bin/env python3
"""
Akna detailide arvutamine tootmisreeglite põhjal.
"""

import pandas as pd
import re

def loe_tootmisreeglid(fail: str = '/Users/makk/Library/CloudStorage/OneDrive-TallinnaTehnikaülikool/6. Semester/Tehisintellekt ja robootika/3. Aknad/Tootmisreeglid.xlsx') -> pd.DataFrame:
    """Loeb tootmisreeglid Exceli failist."""
    return pd.read_excel(fail)

def arvuta_valem(valem: str, laius: int, korgus: int, kogus: int = 1, tkx_korrutis: int = None, tkx_ylal_korrutis: int = None) -> str:
    """Arvutab valemi põhjal väärtuse."""
    if pd.isna(valem) or valem == 'tk':
        return None
    
    # Asenda L ja K vastavate väärtustega
    valem = valem.replace('L', str(laius)).replace('K', str(korgus))
    
    # Arvuta valemi tulemus
    try:
        tulemus = int(eval(valem))
    except:
        return None
    
    # Lisa sulgudesse korrutis, kui on tkx ja tkx_ylal_korrutis on olemas
    if tkx_korrutis is not None and tkx_ylal_korrutis is not None:
        korrutis = tkx_korrutis * tkx_ylal_korrutis * kogus
        return f"{tulemus} ({korrutis}tk)"
    elif tkx_korrutis is not None:
        korrutis = tkx_korrutis * kogus
        return f"{tulemus} ({korrutis}tk)"
    else:
        return tulemus

def arvuta_detailid(df: pd.DataFrame, toote_tyyp: str, laius: int, korgus: int, kasi: str, kogus: int = 1) -> dict:
    """Arvutab akna detailid tootmisreeglite põhjal."""
    detailid = {}
    
    for i in range(len(df)):
        rida = df.iloc[i]
        
        # Kontrolli, kas see on konfigureerimise rida
        if pd.notna(rida['A/U']) and rida['A/U'] in ['A', 'TA']:
            if rida['toote tyyp'] == toote_tyyp and rida['käsi'] == kasi:
                # Arvuta detailid
                detailid['toote_tyyp'] = toote_tyyp
                detailid['laius'] = laius
                detailid['korgus'] = korgus
                detailid['kogus'] = kogus
                detailid['kasi'] = kasi
                
                # Arvuta raami detailid
                if pd.notna(rida['raam']):
                    detailid['raam_laius'] = rida['raam']
                    
                    # Lisa tükkide arv
                    if pd.notna(rida['Unnamed: 10']):
                        tkx = str(rida['Unnamed: 10'])
                        if 'tkx' in tkx:
                            korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                            detailid['raam_laius_tukid'] = korrutis
                
                # Arvuta raami kõrgus
                if pd.notna(rida['Unnamed: 9']):
                    detailid['raam_korgus'] = rida['Unnamed: 9']
                    
                    # Lisa tükkide arv
                    if pd.notna(rida['Unnamed: 13']):
                        tkx = str(rida['Unnamed: 13'])
                        if 'tkx' in tkx:
                            korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                            detailid['raam_korgus_tukid'] = korrutis
                
                # Arvuta klaasi/KLP mõõdud
                if pd.notna(rida['klaasi/KLP mõõt']):
                    detailid['klaasi_laius'] = rida['klaasi/KLP mõõt']
                    
                    # Lisa tükkide arv
                    if pd.notna(rida['Unnamed: 13']):
                        tkx = str(rida['Unnamed: 13'])
                        if 'tkx' in tkx:
                            korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                            detailid['klaasi_laius_tukid'] = korrutis
                
                # Arvuta klaasi/KLP kõrgus
                if pd.notna(rida['Unnamed: 12']):
                    detailid['klaasi_korgus'] = rida['Unnamed: 12']
        
                    # Lisa tükkide arv
                    if pd.notna(rida['Unnamed: 13']):
                        tkx = str(rida['Unnamed: 13'])
                        if 'tkx' in tkx:
                            korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                            detailid['klaasi_korgus_tukid'] = korrutis
                
                # Arvuta järgmise rea detailid
                if i + 1 < len(df):
                    jargmine_rida = df.iloc[i + 1]
                    
                    # Arvuta raami laius
                    if pd.notna(jargmine_rida['raam']):
                        tkx_korrutis = None
                        tkx_ylal_korrutis = None
                        if pd.notna(jargmine_rida['Unnamed: 10']):
                            tkx = str(jargmine_rida['Unnamed: 10'])
                            if 'tkx' in tkx:
                                tkx_korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                        # Extract tkx_ylal_korrutis from the configuration row
                        if i > 0 and pd.notna(df.iloc[i - 1]['Unnamed: 10']) and df.iloc[i - 1]['Unnamed: 10'] != 'tk':
                            tkx_ylal = str(df.iloc[i - 1]['Unnamed: 10'])
                            if tkx_ylal.isdigit():
                                tkx_ylal_korrutis = int(tkx_ylal)
                        detailid['raam_laius_arvutatud'] = arvuta_valem(jargmine_rida['raam'], laius, korgus, kogus, tkx_korrutis, tkx_ylal_korrutis)
                        

                        
                        # Lisa tükkide arv
                        if pd.notna(jargmine_rida['Unnamed: 10']):
                            tkx = str(jargmine_rida['Unnamed: 10'])
                            if 'tkx' in tkx:
                                korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                                detailid['raam_laius_tukid'] = korrutis
                    
                    # Arvuta raami kõrgus
                    if pd.notna(jargmine_rida['Unnamed: 9']):
                        tkx_korrutis = None
                        tkx_ylal_korrutis = None
                        if pd.notna(jargmine_rida['Unnamed: 13']):
                            tkx = str(jargmine_rida['Unnamed: 13'])
                            if 'tkx' in tkx:
                                tkx_korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                        # Extract tkx_ylal_korrutis from the configuration row
                        if i > 0 and pd.notna(df.iloc[i - 1]['Unnamed: 10']) and df.iloc[i - 1]['Unnamed: 10'] != 'tk':
                            tkx_ylal = str(df.iloc[i - 1]['Unnamed: 10'])
                            if tkx_ylal.isdigit():
                                tkx_ylal_korrutis = int(tkx_ylal)
                        detailid['raam_korgus_arvutatud'] = arvuta_valem(jargmine_rida['Unnamed: 9'], laius, korgus, kogus, tkx_korrutis, tkx_ylal_korrutis)
                        

                        
                        # Lisa tükkide arv
                        if pd.notna(jargmine_rida['Unnamed: 13']):
                            tkx = str(jargmine_rida['Unnamed: 13'])
                            if 'tkx' in tkx:
                                korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                                detailid['raam_korgus_tukid'] = korrutis
                    
                    # Arvuta klaasi/KLP mõõdud
                    if pd.notna(jargmine_rida['klaasi/KLP mõõt']):
                        tkx_korrutis = None
                        tkx_ylal_korrutis = None
                        if pd.notna(jargmine_rida['Unnamed: 13']):
                            tkx = str(jargmine_rida['Unnamed: 13'])
                            if 'tkx' in tkx:
                                tkx_korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                        # Extract tkx_ylal_korrutis from the configuration row
                        if i > 0 and pd.notna(df.iloc[i - 1]['Unnamed: 10']) and df.iloc[i - 1]['Unnamed: 10'] != 'tk':
                            tkx_ylal = str(df.iloc[i - 1]['Unnamed: 10'])
                            if tkx_ylal.isdigit():
                                tkx_ylal_korrutis = int(tkx_ylal)
                        detailid['klaasi_laius_arvutatud'] = arvuta_valem(jargmine_rida['klaasi/KLP mõõt'], laius, korgus, kogus, tkx_korrutis, tkx_ylal_korrutis)
                        

                    
                    # Arvuta klaasi/KLP kõrgus
                    if pd.notna(jargmine_rida['Unnamed: 12']):
                        tkx_korrutis = None
                        tkx_ylal_korrutis = None
                        if pd.notna(jargmine_rida['Unnamed: 13']):
                            tkx = str(jargmine_rida['Unnamed: 13'])
                            if 'tkx' in tkx:
                                tkx_korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                        # Extract tkx_ylal_korrutis from the configuration row
                        if i > 0 and pd.notna(df.iloc[i - 1]['Unnamed: 10']) and df.iloc[i - 1]['Unnamed: 10'] != 'tk':
                            tkx_ylal = str(df.iloc[i - 1]['Unnamed: 10'])
                            if tkx_ylal.isdigit():
                                tkx_ylal_korrutis = int(tkx_ylal)
                        detailid['klaasi_korgus_arvutatud'] = arvuta_valem(jargmine_rida['Unnamed: 12'], laius, korgus, kogus, tkx_korrutis, tkx_ylal_korrutis)
                        

                        
                        # Lisa tükkide arv
                        if pd.notna(jargmine_rida['Unnamed: 13']):
                            tkx = str(jargmine_rida['Unnamed: 13'])
                            if 'tkx' in tkx:
                                korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                                detailid['klaasi_korgus_tukid'] = korrutis
                    
                    # Arvuta lengi detailide laius
                    if pd.notna(jargmine_rida['Unnamed: 17']):
                        tkx_korrutis = None
                        tkx_ylal_korrutis = None
                        if pd.notna(jargmine_rida['Unnamed: 16']):
                            tkx = str(jargmine_rida['Unnamed: 16'])
                            if 'tkx' in tkx:
                                tkx_korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                        # Extract tkx_ylal_korrutis from the configuration row
                        if i > 1 and pd.notna(df.iloc[i - 2]['Unnamed: 10']) and df.iloc[i - 2]['Unnamed: 10'] != 'tk':
                            tkx_ylal = str(df.iloc[i - 2]['Unnamed: 10'])
                            if tkx_ylal.isdigit():
                                tkx_ylal_korrutis = int(tkx_ylal)
                        detailid['lengi_laius_arvutatud'] = arvuta_valem(jargmine_rida['Unnamed: 17'], laius, korgus, kogus, tkx_korrutis, tkx_ylal_korrutis)
                        

                    
                    # Arvuta lengi detailide kõrgus
                    if pd.notna(jargmine_rida['Unnamed: 19']):
                        tkx_korrutis = None
                        tkx_ylal_korrutis = None
                        if pd.notna(jargmine_rida['Unnamed: 20']):
                            tkx = str(jargmine_rida['Unnamed: 20'])
                            if 'tkx' in tkx:
                                tkx_korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                        # Extract tkx_ylal_korrutis from the configuration row
                        if i > 0 and pd.notna(df.iloc[i - 1]['Unnamed: 10']) and df.iloc[i - 1]['Unnamed: 10'] != 'tk':
                            tkx_ylal = str(df.iloc[i - 1]['Unnamed: 10'])
                            if tkx_ylal.isdigit():
                                tkx_ylal_korrutis = int(tkx_ylal)
                        else:
                            tkx_ylal_korrutis = 1  # Default value if not found
                        detailid['lengi_korgus_arvutatud'] = arvuta_valem(jargmine_rida['Unnamed: 19'], laius, korgus, kogus, tkx_korrutis, tkx_ylal_korrutis)
                        

                    
                    # Arvuta raami detailid
                    if pd.notna(jargmine_rida['raami detailid']):
                        tkx_korrutis = None
                        tkx_ylal_korrutis = None
                        if pd.notna(jargmine_rida['Unnamed: 22']):
                            tkx = str(jargmine_rida['Unnamed: 22'])
                            if 'tkx' in tkx:
                                tkx_korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                        if pd.notna(jargmine_rida['Unnamed: 20']):
                            tkx_ylal = str(jargmine_rida['Unnamed: 20'])
                            if tkx_ylal.isdigit():
                                tkx_ylal_korrutis = int(tkx_ylal)
                        detailid['raami_detailid_arvutatud'] = arvuta_valem(jargmine_rida['raami detailid'], laius, korgus, kogus, tkx_korrutis, tkx_ylal_korrutis)
                        

                    
                    # Arvuta raami liistude laius
                    if pd.notna(jargmine_rida['Unnamed: 23']):
                        tkx_korrutis = None
                        tkx_ylal_korrutis = None
                        if pd.notna(jargmine_rida['Unnamed: 24']):
                            tkx = str(jargmine_rida['Unnamed: 24'])
                            if 'tkx' in tkx:
                                tkx_korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                        if pd.notna(jargmine_rida['Unnamed: 23']):
                            tkx_ylal = str(jargmine_rida['Unnamed: 23'])
                            if tkx_ylal.isdigit():
                                tkx_ylal_korrutis = int(tkx_ylal)
                        detailid['raami_liistud_laius_arvutatud'] = arvuta_valem(jargmine_rida['Unnamed: 23'], laius, korgus, kogus, tkx_korrutis, tkx_ylal_korrutis)
                        

                    
                    # Arvuta raami liistude kõrgus
                    if pd.notna(jargmine_rida['Unnamed: 25']):
                        tkx_korrutis = None
                        tkx_ylal_korrutis = None
                        if pd.notna(jargmine_rida['Unnamed: 26']):
                            tkx = str(jargmine_rida['Unnamed: 26'])
                            if 'tkx' in tkx:
                                tkx_korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                        if pd.notna(jargmine_rida['Unnamed: 25']):
                            tkx_ylal = str(jargmine_rida['Unnamed: 25'])
                            if tkx_ylal.isdigit():
                                tkx_ylal_korrutis = int(tkx_ylal)
                        detailid['raami_liistud_korgus_arvutatud'] = arvuta_valem(jargmine_rida['Unnamed: 25'], laius, korgus, kogus, tkx_korrutis, tkx_ylal_korrutis)
                        

                    
                    # Arvuta klaasiliistud
                    if pd.notna(jargmine_rida['klaasiliistud']):
                        tkx_korrutis = None
                        tkx_ylal_korrutis = None
                        if pd.notna(jargmine_rida['Unnamed: 28']):
                            tkx = str(jargmine_rida['Unnamed: 28'])
                            if 'tkx' in tkx:
                                tkx_korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                        if pd.notna(jargmine_rida['Unnamed: 26']):
                            tkx_ylal = str(jargmine_rida['Unnamed: 26'])
                            if tkx_ylal.isdigit():
                                tkx_ylal_korrutis = int(tkx_ylal)
                        detailid['klaasiliistud_arvutatud'] = arvuta_valem(jargmine_rida['klaasiliistud'], laius, korgus, kogus, tkx_korrutis, tkx_ylal_korrutis)
                        

    
    return detailid

def main():
    """Peamine funktsioon kasutajaliidesega."""
    print("Akna detailide arvutamine")
    print("=" * 50)
    
    akna_tyyp = input("Sisesta akna tüüp (tavaline/topelt): ").lower()
    toote_tyyp = input("Sisesta toote tüüp (nt 40STAND60x63, 77STAND60x70): ")
    laius = int(input("Sisesta laius (mm): "))
    korgus = int(input("Sisesta kõrgus (mm): "))
    kogus = int(input("Sisesta akende kogus: "))
    kasi = input("Sisesta käelisus (P/V): ").upper()
    
    df = loe_tootmisreeglid()
    detailid = arvuta_detailid(df, toote_tyyp, laius, korgus, kasi, kogus)
    
    print("\nArvutused detailid:")
    print("=" * 50)
    print(f"Akende kogus: {kogus}")
    for key, value in detailid.items():
        print(f"{key}: {value}")

if __name__ == '__main__':
    main()