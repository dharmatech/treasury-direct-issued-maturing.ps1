#!/bin/sh

cd /var/www/dharmatech.dev/data/treasury-direct-issued-maturing.ps1

mkdir -p ../reports/treasury-direct-issued-maturing

# pwsh treasury-direct-issued-maturing.ps1 > ../reports/treasury-direct-issued-maturing/`date +%Y-%m-%d`.txt


# pwsh -Command "./treasury-direct-issued-maturing.ps1 | Out-String > ../reports/treasury-direct-issued-maturing/`date +%Y-%m-%d`.txt"

# pwsh -Command "./treasury-direct-issued-maturing.ps1 | Out-String *> ../reports/treasury-direct-issued-maturing/`date +%Y-%m-%d`.txt"

# pwsh -Command "./treasury-direct-issued-maturing.ps1 *>&1 | Out-String > ../reports/treasury-direct-issued-maturing/`date +%Y-%m-%d`.txt"

# pwsh -Command "./treasury-direct-issued-maturing.ps1 *>&1 | Out-String -Stream -Width 230 > ../reports/treasury-direct-issued-maturing/`date +%Y-%m-%d`.txt"

pwsh -Command "./treasury-direct-issued-maturing.ps1 *>&1 | Out-String -Stream -Width 230 > script-out"

cat script-out | 
    /home/dharmatech/go/bin/terminal-to-html -preview |
    sed 's/pre-wrap/pre/' |
    sed 's/terminal-to-html Preview/treasury-direct-issued-maturing.ps1/' |
    sed 's/<body>/<body style="width: fit-content;">/' > ../reports/treasury-direct-issued-maturing/`date +%Y-%m-%d`.html
