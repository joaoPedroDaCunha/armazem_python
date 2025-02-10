if checkbox_lote2_var.get() == 1:
    if nf1 == nf2 :
        ws_descarga_sal['D20'] = nf1
        ws_descarga_sal['K20'] = nfpalete1
        ws_descarga_sal['P20'] = peso1+peso2
        ws_descarga_sal['D22'] = " "
        ws_descarga_sal['K22'] = " "
        ws_descarga_sal['P22'] = " "
        if checkbox_lote3_var.get() == 1 :
            ws_descarga_sal['D24'] = nf3
            ws_descarga_sal['K24'] = nfpalete3
            ws_descarga_sal['P24'] = peso3
        else :
            ws_descarga_sal['D24'] = " "
            ws_descarga_sal['K24'] = " "
            ws_descarga_sal['P24'] = " "
elif checkbox_lote3_var.get() == 1 :
    if nf1 == nf2 and nf1 == nf3 :
        ws_descarga_sal['D20'] = nf1
        ws_descarga_sal['K20'] = nfpalete1
        ws_descarga_sal['P20'] = peso1+peso2+peso3
        ws_descarga_sal['D22'] = " "
        ws_descarga_sal['K22'] = " "
        ws_descarga_sal['P22'] = " "
        ws_descarga_sal['D24'] = " "
        ws_descarga_sal['K24'] = " "
        ws_descarga_sal['P24'] = " "


    ws_descarga_sal['D20'] = nf1
    ws_descarga_sal['K20'] = nfpalete1
    ws_descarga_sal['P20'] = peso1
    if checkbox_lote2_var.get() == 1:
        if nf1 == nf2 :
            ws_descarga_sal['D22'] = " "
            ws_descarga_sal['K22'] = " "
            ws_descarga_sal['P22'] = " "
            ws_descarga_sal['P20'] = peso1+peso2
        else :
            ws_descarga_sal['D22'] = nf2
            ws_descarga_sal['K22'] = nfpalete2
            ws_descarga_sal['P22'] = peso2
    else :
        ws_descarga_sal['D22'] = " "
        ws_descarga_sal['K22'] = " "
        ws_descarga_sal['P22'] = " "
    if checkbox_lote3_var.get() == 1:
        if nf1 == nf3 and nf1 == nf2:
            ws_descarga_sal['D24'] = " "
            ws_descarga_sal['K24'] = " "
            ws_descarga_sal['P24'] = " "
            ws_descarga_sal['P20'] = peso1+peso2+peso3
        else:
            ws_descarga_sal['D24'] = nf3
            ws_descarga_sal['K24'] = nfpalete3
            ws_descarga_sal['P24'] = peso3
    else :
        ws_descarga_sal['D24'] = " "
        ws_descarga_sal['K24'] = " "
        ws_descarga_sal['P24'] = " "