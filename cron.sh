#!/bin/sh

cd /var/www/dharmatech.dev/data/treasury-direct-issued-maturing.ps1

# screen -d -m ./to-report.sh

# tmux new-session -d -x 300 bash -c 'script -q -c "pwsh ./format-table-example.ps1"'

tmux new-session -d -x 300 bash -c 'script -q -c ./to-report.sh'
