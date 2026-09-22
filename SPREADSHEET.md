# Wiring the roster spreadsheet to Warcalc

Four codes travel between the sheet and the calculator.

| Code | What it is | Who writes it |
|---|---|---|
| **army code** | one side's units — type, weapon, armour, size, training for every stack | Warcalc's copy button, pasted into the sheet by hand |
| **crunch** | the 43 roster counts, and nothing else | a formula in the sheet, always live |
| **battle code** | both army codes, the defender bonus and the crunch, in one string | a formula in the sheet |
| **link** | the battle code with the site in front of it | a formula in the sheet |

The last two are the same string; the link just prefixes it. Hand out whichever
suits the reader — the link for anyone on the site, the battle code for anyone
running their own copy of the page. Both carry the crunch, so both get checked.

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

Counts live in **C10:C52** — Cavalry Amount in C10, Alicorns Amount in C52 —
in the roster's own row order (the table at the bottom). The letter comes from
the row's offset within the range, not its absolute row, so if you move the
block just change both references to match its new first cell.

**Google Sheets**

```
=TEXTJOIN("",TRUE,ARRAYFORMULA(IF(ROUND(C10:C52,0)>0,MID("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopq",ROW(C10:C52)-ROW(C$10)+1,1)&ROUND(C10:C52,0),"")))
```

**Excel 365**

```
=TEXTJOIN("",TRUE,IF(ROUND(C10:C52,0)>0,MID("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopq",SEQUENCE(43),1)&ROUND(C10:C52,0),""))
```

**Excel 2016 / 2019** — no `SEQUENCE`, and `IF` over a range needs a helper
column. Put this in **D10** and fill down to **D52**:

```
=IF(ROUND(C10,0)>0,MID("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopq",ROW()-ROW(C$10)+1,1)&ROUND(C10,0),"")
```

then collapse it:

```
=TEXTJOIN("",TRUE,D10:D52)
```

(`=CONCAT(D10:D52)` works too. On anything older than 2016, `=D10&D11&D12&…`.)

---

## 2. The battle code, and the link

Two things come out of the sheet, and they carry the same battle:

- **the battle code** — paste it into Warcalc's **Import battle** box
- **the link** — the same code with the site in front of it, so clicking loads it

The crunch rides along inside both. It is not a separate thing to hand out; it
is just how the roster's numbers reach Warcalc so that Warcalc can do the
comparison. Whoever opens the battle gets the mismatch report either way.

Fill in three cells by hand:

| Cell | Holds |
|---|---|
| `C53` | defender bonus — `Yes` or `No` |
| `C55` | attacker army code, pasted from Warcalc |
| `C56` | defender army code, pasted from Warcalc |

Then three formulas:

**C58 — the crunch.** Plumbing; nobody reads this cell.

```
=TEXTJOIN("",TRUE,ARRAYFORMULA(IF(ROUND(C10:C52,0)>0,MID("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopq",ROW(C10:C52)-ROW(C$10)+1,1)&ROUND(C10:C52,0),"")))
```

**C60 — the battle code.**

```
=IF(AND(C55="",C56=""),"","+"&C55&IF(C53="No","=","-")&C56&IF(C58="","","*"&C58))
```

```
+AFd-0*b3d3g3
```

**C57 — the link.**

```
=IF(C60="","",HYPERLINK("https://cloppenmanager.neocities.org/warcalc#"&C60,"Open in Warcalc"))
```

For the URL as plain text instead of a live link:

```
=IF(C60="","","https://cloppenmanager.neocities.org/warcalc#"&C60)
```

```
https://cloppenmanager.neocities.org/warcalc#+AFd-0*b3d3g3
```

### What Warcalc does with it

Opening the link, or pasting the battle code into **Import battle**, loads both
sides and the defender bonus. Warcalc then adds up its own 43 totals from the
army codes it just loaded and compares them against the crunch that came in
with them. If they disagree it puts a report over the board naming every row
that is off, which side it is on, what the sheet said and what the code
actually carried. If they agree it says nothing.

The battle loads either way — the point is to catch the drift, not to refuse
the fight. And the link is read once and then dropped from the address bar, so
reloading keeps whatever the player has since changed.

### If your roster block is the defenders

`C58` is treated as the **attackers'** roster. If the block at C10:C52 is
really your defenders, change the tail of `C60` from `"*"&C58` to `"**"&C58`
and it will be checked against that side instead.

### Grammar

```
[<url>#]<battle code>[*<attacker crunch>[*<defender crunch>]]
```

`+` opens the attackers. `-` opens the defenders with the defender bonus on,
`=` with it off. Either crunch may be left empty. **Import battle** accepts any
of it — the bare battle code, the code with crunches, or the whole URL — so it
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
