# Aknadetailide arvutusrakendus

## Ülesande kirjeldus

Loo Pythonis CLI-rakendus, mis arvutab akende tootmiseks vajalikud mõõdud Exceli tabeli valemite põhjal.

## Exceli tabeli struktuur

### Sisendtulbad (kliendi antavad mõõdud):

| Tulp | Kirjeldus | Väärtused |
|------|-----------|-----------|
| A | Aknatüüp | A = tavaline aken, TA = topeltaken |
| B | Tootetüüp | - |
| C | Laius (L) | Millimeetrites |
| D | Kõrgus (K) | Millimeetrites |
| E | Kogus | Mitu akent |
| F | Käelisus | P = parempoolne, V = vasakpoolne |

### Valemitega tulbad (arvutatavad väärtused):

| Tulbad | Kirjeldus |
|--------|-----------|
| G-I | Raami mõõtmed (sõltuvad L ja K väärtustest) |
| J-L | Klaasi mõõtmed (sõltuvad L ja K väärtustest) |
| P | Lengi detail 1 |
| R | Lengi detail 2 |
| T-Y | Raami detailide pikkused ja kogused |
| Z-AC | Klaasiliistude pikkused ja kogused |

## Rakenduse nõuded

1. **Exceli lugemine**: Kasuta `pandas` või `openpyxl`.
2. **Kasutaja sisend**: Küsib terminalist:
   - Aknatüüp (A/TA)
   - Tootetüüp
   - Laius (L)
   - Kõrgus (K)
   - Kogus
   - Käelisus (P/V)
3. **Valemite arvutus**: Exceli valemid peavad arvutama mõõdud dünaamiliselt.
4. **Väljund**: Kuvab terminalis:
   - Raami mõõdud
   - Klaasi mõõdud
   - Lengi detailid
   - Raami detailid
   - Klaasiliistud
   - Kogus × sisestatud kogus

## Näide

### Sisend:
```
Aknatüüp: A
Toote tüüp: 2
Laius (L): 800
Kõrgus (K): 1200
Kogus: 2
Käelisus: P
```

### Väljund:
```
Raami mõõdud:
 - Raam L1: 756 mm
 - Raam L2: 1156 mm

Klaasi mõõdud:
 - Klaas L: 700 mm
 - Klaas K: 1100 mm

Raami detailid:
 - Ülemine detail: 756 mm
 - Alumine detail: 756 mm
 - Külgdetailid: 1156 mm

Klaasiliistud:
 - Liist 1: 700 mm
 - Liist 2: 1100 mm

Kogus: 2
```

## Tehnilised nõuded

- **Python 3**: Rakendus peab töötama Python 3 all.
- **Dünaamilised valemid**: Exceli valemid peavad arvutama mõõdud dünaamiliselt.
- **Kohanduvat**: Kui Exceli tabelit muudetakse, peab programm kasutama uusi väärtusi.
- **Kasutajasõbralik**: Väljund peab olema struktureeritud ja loetav.

## Lõppeesmärk

Loo Python CLI-rakendus, mis:
1. Küsib kasutajalt sisendmõõdud.
2. Arvutab Exceli valemite põhjal kõik aknadetailid.
3. Kuvab tulemuse terminalis.
4. Kohandub automaatselt Exceli tabeli muudatustega.
5. Sulgudes peale arvutusi näha vajaminevate detailide arvud
