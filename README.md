## FLPL-checker
Flat steel stock checker for Tekla.

Compares flat profiles in Tekla model with .csv stock list. Changes prefix of profiles to FL or PL accordingly.

Built for Skanding to use with DS stock system.

### Installation guide
Built as a Tekla macro. Copy file 'CodeFile1.cs' (found in folder Project1) to your 
> Tekla installation folder -> Environments -> Common -> macros -> modelling

Create 'stock list.csv' and point to it in the 'CodeFile1.cs' (line 25).
