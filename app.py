def detect_invoice_fields(text: str, filename: str):
    scanned = looks_scanned(text)

    # PROVEEDOR
    proveedor = find_first(
        [
            r"NAVATEC\s+INGENIERIA\s+S\.A\.",
            r"NAVATECO",
        ],
        text,
    )

    # NIT
    nit = find_first(
        [
            r"Ident\.\s*Jur[ií]dica:\s*([0-9\-]+)",
        ],
        text,
    )

    # Número factura
    factura_num = find_first(
        [
            r"Factura\s+Electr[oó]nica\s+N°\s*([0-9]+)",
        ],
        text,
    )

    # Fecha
    fecha = find_first(
        [
            r"Fecha\s+de\s+Emisi[oó]n:\s*([0-9\/\:\sa\.m\.p\.]+)",
        ],
        text,
    )

    # Subtotal
    subtotal_str = find_first(
        [
            r"Subtotal\s+Neto\s*¢\s*([0-9\.,]+)",
        ],
        text,
    )
    subtotal = parse_number_co(subtotal_str)

    # IVA
    iva_str = find_first(
        [
            r"Total\s+Impuesto\s*¢\s*([0-9\.,]+)",
        ],
        text,
    )
    iva = parse_number_co(iva_str)

    # Total
    total_str = find_first(
        [
            r"Total\s+Factura:\s*¢\s*([0-9\.,]+)",
        ],
        text,
    )
    total = parse_number_co(total_str)

    return InvoiceExtract(
        documento=filename,
        es_probable_escaneado=scanned,
        proveedor=proveedor,
        nit=nit,
        factura_numero=factura_num,
        fecha=fecha,
        subtotal=subtotal,
        iva=iva,
        total=total,
        moneda="CRC",
        confidence_hint="Reglas NAVATEC",
    )
