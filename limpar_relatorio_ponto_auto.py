# -*- coding: utf-8 -*-
import os
import sys
import subprocess
from pathlib import Path

def main():
    pasta = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent

    entrada1 = pasta / "RELATORIO BRUTO.csv"
    entrada2 = pasta / "RELATÓRIO BRUTO.csv"

    if entrada1.exists():
        entrada = entrada1
    elif entrada2.exists():
        entrada = entrada2
    else:
        print("ERRO: Arquivo bruto não encontrado.")
        print("Coloque o arquivo com um destes nomes na mesma pasta do programa:")
        print("- RELATORIO BRUTO.csv")
        print("- RELATÓRIO BRUTO.csv")
        input("Pressione ENTER para sair...")
        return

    saida = pasta / "RELATORIO LIMPO.csv"

    print("Arquivo encontrado:", entrada.name)
    print("Gerando relatório limpo...")
    print()

    try:
        # Importa o script principal
        from limpar_relatorio_ponto import limpar_relatorio

        total, caminho_saida = limpar_relatorio(
            input_path=str(entrada),
            output_path=str(saida),
            max_marcacoes=12
        )

        print()
        print("CONCLUÍDO COM SUCESSO!")
        print("Arquivo gerado:", saida.name)
        print("Linhas geradas:", total)
        print()

    except Exception as e:
        print()
        print("ERRO AO PROCESSAR O ARQUIVO:")
        print(str(e))
        print()

    input("Pressione ENTER para sair...")

if __name__ == "__main__":
    main()
