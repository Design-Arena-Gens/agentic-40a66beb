"""Command line interface for exporting Italian companies by ATECO code."""

from __future__ import annotations

import argparse
import os
import sys
from typing import Iterable, Sequence

from .client import (
    OpenAPICompanyClient,
    OpenAPIError,
    sanitize_ateco_code,
)
from .exporter import export_to_excel


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Estrai l'elenco delle aziende italiane in base al codice ATECO "
            "e provincia utilizzando la API Company di OpenAPI e salva "
            "i risultati in formato XLSX."
        )
    )

    parser.add_argument(
        "--ateco",
        required=True,
        help="Codice ATECO primario da cercare (es. 1071 o 10.71).",
    )
    parser.add_argument(
        "--province",
        default="VR",
        help="Provincia (codice di due lettere, default: VR).",
    )
    parser.add_argument(
        "--token",
        default=None,
        help=(
            "Token di accesso OpenAPI. Se non fornito viene letto dalla "
            "variabile d'ambiente OPENAPI_TOKEN."
        ),
    )
    parser.add_argument(
        "--output",
        default="companies.xlsx",
        help="Percorso del file XLSX di destinazione (default: companies.xlsx).",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=100,
        help="Numero massimo di record per pagina (max 100).",
    )
    parser.add_argument(
        "--max-records",
        type=int,
        default=None,
        help="Limite massimo di record totali da scaricare (default: tutti).",
    )
    parser.add_argument(
        "--activity-status",
        default=None,
        help="Filtra per stato attività (es. ATTIVA).",
    )
    parser.add_argument(
        "--data-enrichment",
        default="Advanced",
        choices=[
            "",
            "Start",
            "Advanced",
            "Address",
            "Pec",
            "Shareholders",
        ],
        help="Dataset di arricchimento da includere (default: Advanced).",
    )
    parser.add_argument(
        "--sandbox",
        action="store_true",
        help="Usa l'ambiente sandbox di test.company.openapi.com",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Non scarica i dati, restituisce solo il numero di risultati disponibili.",
    )

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    token = args.token or os.environ.get("OPENAPI_TOKEN")
    if not token:
        parser.error(
            "È necessario fornire un token OpenAPI tramite --token o la variabile "
            "d'ambiente OPENAPI_TOKEN."
        )

    try:
        sanitized_ateco = sanitize_ateco_code(args.ateco)
    except ValueError as exc:
        parser.error(str(exc))

    client = OpenAPICompanyClient(token, sandbox=args.sandbox)

    if args.dry_run:
        try:
            total = client.dry_run_count(
                province=args.province,
                ateco_code=sanitized_ateco,
                activity_status=args.activity_status,
            )
        except OpenAPIError as exc:  # pragma: no cover - thin wrapper around API
            print(f"Errore durante la dry run: {exc}", file=sys.stderr)
            return 1

        print(
            f"Sono disponibili {total} aziende per ATECO {sanitized_ateco} in provincia {args.province.upper()}."
        )
        return 0

    try:
        records: Iterable[dict] = client.search_companies(
            province=args.province,
            ateco_code=sanitized_ateco,
            data_enrichment=args.data_enrichment,
            limit=args.limit,
            max_records=args.max_records,
            activity_status=args.activity_status,
        )
        count = export_to_excel(records, args.output)
    except OpenAPIError as exc:  # pragma: no cover - thin wrapper
        print(f"Errore durante l'estrazione: {exc}", file=sys.stderr)
        return 1

    print(
        f"Salvate {count} aziende con codice ATECO {sanitized_ateco} in provincia {args.province.upper()} "
        f"nel file '{args.output}'."
    )
    return 0


if __name__ == "__main__":  # pragma: no cover - entry point
    sys.exit(main())
