# Wiring the roster spreadsheet to Warcalc

Four codes travel between the sheet and the calculator.

| Code | What it is | Who writes it |
|---|---|---|
| **army code** | one side's units — type, weapon, armour, size, training for every stack | Warcalc's copy button, pasted into the sheet by hand |
| **crunch** | the 43 roster counts, and nothing else | a formula in the sheet, always live |
| **battle code** | both army codes, the defender bonus, and optionally both crunches | a formula in the sheet |
| **link** | the battle code with the site in front of it | a formula in the sheet |

The last two are the same string; the link just prefixes it. Hand out whichever
suits the reader — the link for anyone on the site, the battle code for anyone
running their own copy of the page.

The army code is the only thing a human has to keep up to date, and it is the
only thing that can go stale. That is what the crunch is for: it is computed
from the roster cells every time they change, so when the two disagree Warcalc
can say which rows are wrong.

**The sheet cannot produce the army code on its own.** It counts stock — how
many unicorns, how many Long Weights, how many Librarians — but never records
which weapon sits next to which armour. Three unicorns with Long Weight +
Librarian and three unicorns split two/one across two loadouts give identical
roster rows. The army code needs the pairing, so it still gets pasted in from
Warcalc; the crunch just makes sure nobody forgets to.

---

## 1. The crunch

Slot `i` is written as its letter followed by the count in plain decimal.
Letters and digits never collide, so nothing needs a separator, and zero rows
are left out entirely. Three unicorns carrying Long Weights behind Librarians
is:

```
b3d3g3
```

### The formula

Counts live in **C1:C43**, in the roster's own row order (the table at the
bottom). Adjust the range if your block starts somewhere else — the letter is
worked out from the row's offset within the range, not from its absolute row.

**Google Sheets**

```
=TEXTJOIN("",TRUE,ARRAYFORMULA(IF(ROUND(C1:C43,0)>0,
  MID("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopq",ROW(C1:C43)-ROW(C$1)+1,1)
  &ROUND(C1:C43,0),"")))
```

**Excel 365**

```
=TEXTJOIN("",TRUE,IF(ROUND(C1:C43,0)>0,
  MID("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopq",SEQUENCE(43),1)
  &ROUND(C1:C43,0),""))
```

**Excel 2016 / 2019** — no `SEQUENCE`, and `IF` over a range needs a helper
column. Put this in **D1** and fill down to **D43**:

```
=IF(ROUND(C1,0)>0,MID("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopq",ROW()-ROW(C$1)+1,1)&ROUND(C1,0),"")
```

then collapse it:

```
=TEXTJOIN("",TRUE,D1:D43)
```

(`=CONCAT(D1:D43)` works too. On anything older than 2016, `=D1&D2&D3&…`.)

---

## 2. The battle code, and the link

Everything composes out of one cell. Build the battle code first, hang the
roster check off it, and the link is just the site in front of both.

A suggested layout — move the cells wherever you like, the formulas only care
about the references:

| Cell | Holds |
|---|---|
| `$B$1` | `https://cloppenmanager.neocities.org/warcalc` |
| `C53` | defender bonus — `Yes` or `No` |
| `C55` | attacker army code, pasted from Warcalc |
| `C56` | defender army code, pasted from Warcalc |
| `C58` | attacker crunch — the formula from part 1 |
| `C59` | defender crunch — the same formula over the defenders' block |

**C60 — the battle code, on its own.** For anyone running their own copy of
the page who just wants to paste something into **Import battle**:

```
=IF(AND(C55="",C56=""),"","+"&C55&IF(C53="No","=","-")&C56)
```

```
+AFd-0
```

**C61 — the battle code with the roster check on the end.** Same thing, plus
whichever crunches exist. This is the one worth handing out: **Import battle**
takes it exactly like the plain code, and it still catches a stale roster:

```
=IF(C60="","",C60&IF(AND(C58="",C59=""),"","*"&C58&IF(C59="","","*"&C59)))
```

```
+AFd-0*b3d3g3
```

**C57 — the clickable link.** Same formula in Excel and Sheets:

```
=IF(C61="","",HYPERLINK($B$1&"#"&C61,"Open in Warcalc"))
```

If you want the URL as plain text you can copy rather than a live link, drop
the `HYPERLINK` wrapper:

```
=IF(C61="","",$B$1&"#"&C61)
```

```
https://cloppenmanager.neocities.org/warcalc#+AFd-0*b3d3g3
```

So there are three things to hand out, and all three carry the same battle:

| Cell | Give it to | Loads by |
|---|---|---|
| `C57` | anyone on the site | clicking it |
| `C61` | anyone on a local copy | pasting into **Import battle** |
| `C60` | anyone who does not want the roster check | pasting into **Import battle** |

Clicking the link loads both sides and the defender bonus. If a crunch
disagrees with the army code it came in with, Warcalc loads the battle anyway
and puts a report on top of it naming every row that is off, which side it is
on, what the sheet said and what the code actually carried. A crunch that
agrees is silent. Pasting `C61` into **Import battle** does the same thing,
report and all.

The link is read once and then dropped from the address bar, so reloading keeps
whatever the player has since changed instead of snapping back to the link.

### Grammar

```
[<url>#]<battle code>[*<attacker crunch>[*<defender crunch>]]
```

`+` opens the attackers. `-` opens the defenders with the defender bonus on,
`=` with it off. Either crunch may be left empty — `*` `*` `<def>` sends only
the defenders' one. Everything from the battle code rightward is what
**Import battle** accepts; it will also take the whole URL, `#` and all, so it
does not matter which cell someone copied.

## 3. Roster rows and their letters

| 1 | Cavalry Amount | `A` |
| 2 | PRC-E6 | `B` |
| 3 | PRC-E7 | `C` |
| 4 | PRC-E8 | `D` |
| 5 | ACFU | `E` |
| 6 | APFU | `F` |
| 7 | ATFU | `G` |
| 8 | AUFU | `H` |
| 9 | Barding | `I` |
| 10 | Bigdog | `J` |
| 11 | Nope | `K` |
| 12 | Tanks Amount | `L` |
| 13 | K9P | `M` |
| 14 | ELBO-GRS | `N` |
| 15 | Chem-Light Battery | `O` |
| 16 | Trundle | `P` |
| 17 | Shepherd | `Q` |
| 18 | Ohno | `R` |
| 19 | Titan | `S` |
| 20 | Pegasi Amount | `T` |
| 21 | Prop Wash | `U` |
| 22 | Steam Bucket | `V` |
| 23 | Canopy Lights | `W` |
| 24 | Cooler | `X` |
| 25 | Wonder | `Y` |
| 26 | Griffin | `Z` |
| 27 | Dragon | `a` |
| 28 | Unicorns Amount | `b` |
| 29 | Long Stand | `c` |
| 30 | Long Weight | `d` |
| 31 | Grid Squares | `e` |
| 32 | Hornshield | `f` |
| 33 | Librarian | `g` |
| 34 | Shining | `h` |
| 35 | D2A | `i` |
| 36 | Naval Amount | `j` |
| 37 | Shoreline | `k` |
| 38 | Water Hammer | `l` |
| 39 | Waterline Eraser | `m` |
| 40 | C-Pon3 | `n` |
| 41 | Esohes | `o` |
| 42 | Shubidu | `p` |
| 43 | Alicorns Amount | `q` |

Rows the sheet does not track — Scrounged Weapons and Scrounged Armor — have no
letter. A scrounged unit still counts toward its type's **Amount** row and
toward nothing else, which is what the sheet does too.

Note that the cavalry block lists **APFU before ATFU**, which is the opposite of
the order inside Warcalc's own dropdowns. The table above is the one that
matters; do not re-sort it.
