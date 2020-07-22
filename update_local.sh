#!/bin/bash

gdrive='/home/rip/gebdiwi/'

python3 odf_analyse.py --directory ${gdrive}Gebiete/Gruppe\ 0/ods/ \
${gdrive}Gebiete/Gruppe\ 1/ods/ \
${gdrive}Gebiete/Gruppe\ 2/ods/ \
${gdrive}Gebiete/Gruppe\ 3/ods/ \
--refresh --overview ${gdrive}/Gebiete/Gebietsuebersicht.ods

exit 0
