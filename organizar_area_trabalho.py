#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ORGANIZADOR DE ARQUIVOS DA ÁREA DE TRABALHO
Lê os nomes/caminhos dos arquivos, classifica-os por assunto e os reorganiza
em pastas com nomes em português.

SEGURANÇA:
- Começa apenas em modo de simulação; nada é movido antes da confirmação.
- Não sobrescreve arquivos com nomes iguais.
- Preserva integralmente pastas com "lagora", "censo" ou "sistema presença"
  no nome (inclusive tudo o que estiver dentro delas).
- Nunca move nenhum arquivo chamado "shutdown.bat" e preserva as pastas
  ancestrais necessárias para que ele continue no mesmo lugar.
- Só remove pastas que ficarem vazias. Nunca apaga arquivos.
- Cria um registro CSV de origem/destino e um log de execução.

Uso: python organizar_area_trabalho.py
"""

from pathlib import Path
from datetime import datetime
import csv
import os
import re
import shutil
import sys

def encontrar_area_de_trabalho():
    """Localiza a Área de Trabalho do Windows, inclusive quando redirecionada."""
    script_dir = Path(__file__).resolve().parent
    candidatos = [
        Path.home() / "Desktop",
        Path.home() / "Área de Trabalho",
        Path.home() / "OneDrive" / "Desktop",
        Path.home() / "OneDrive" / "Área de Trabalho",
    ]
    # Se o script já estiver numa pasta chamada Desktop, use exatamente ela.
    if script_dir.name.casefold() in {"desktop", "área de trabalho"}:
        return script_dir
    for candidato in candidatos:
        if candidato.is_dir():
            return candidato
    # Último recurso: operar na pasta do script, mas avisar no início.
    return script_dir

RAIZ = encontrar_area_de_trabalho()

# Arquivos auxiliares do organizador não serão movidos durante a execução.
NOME_SCRIPT = Path(__file__).name
NOME_RELATORIO_INVENTARIO = "folder_inventory.txt"
NOME_LOG = "registro_organizacao.csv"

# Pastas protegidas: não mover, renomear, esvaziar nem alterar seu conteúdo.
# A comparação ignora maiúsculas, acentos e separadores comuns.
TERMOS_PASTAS_PROTEGIDAS = ("lagora", "censo", "sistema presenca")
ARQUIVO_PROTEGIDO = "shutdown.bat"

PASTA_DESTINO_GERAL = {
    "escola": "Escola - Secretaria",
    "frequencia": "Escola - Frequência e Presença",
    "atas": "Escola - Atas e Resultados",
    "censo": "Escola - Censo Escolar",
    "alunos": "Escola - Alunos e Responsáveis",
    "atestados": "Escola - Atestados e Declarações",
    "inclusao": "Escola - Inclusão e Atendimento",
    "passeio": "Escola - Passeios e Autorizações",
    "biblioteca": "Escola - Biblioteca e Livros",
    "financeiro": "Escola - Financeiro e Prestação de Contas",
    "trabalho": "Estudos - Faculdade e Matemática",
    "programacao": "Programação e Projetos",
    "instaladores": "Programas - Instaladores e Sistemas",
    "imagens": "Imagens e Fotografias",
    "midia": "Áudio e Vídeo",
    "compactados": "Arquivos Compactados",
    "documentos": "Documentos Gerais",
    "pdf": "Documentos - PDFs",
    "word": "Documentos - Word e Texto",
    "planilhas": "Planilhas e Dados",
    "dados": "Dados e Páginas Web",
    "temporarios": "Temporários e Arquivos Auxiliares",
    "organizacao": "Ferramentas de Organização",
}

EXTENSOES_IMAGEM = {
    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tif",
    ".tiff", ".ico", ".heic", ".svg"
}
EXTENSOES_AUDIO_VIDEO = {
    ".mp3", ".wav", ".ogg", ".flac", ".m4a", ".aac", ".mp4",
    ".mkv", ".avi", ".mov", ".webm", ".wmv"
}
EXTENSOES_COMPACTADOS = {".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz"}
EXTENSOES_INSTALADORES = {
    ".exe", ".msi", ".iso", ".img", ".dmg", ".deb", ".rpm", ".appx",
    ".msix", ".dll", ".sys", ".cab"
}
EXTENSOES_CODIGO = {
    ".py", ".pyw", ".js", ".jsx", ".ts", ".tsx", ".html", ".htm",
    ".css", ".json", ".xml", ".yaml", ".yml", ".toml", ".sh", ".bat",
    ".cmd", ".ps1", ".sql", ".java", ".c", ".h", ".cpp", ".hpp",
    ".cs", ".go", ".rs", ".php", ".rb", ".r", ".ipynb", ".vue",
    ".svelte", ".md", ".rst"
}
EXTENSOES_TEMPORARIOS = {".tmp", ".temp", ".bak", ".old", ".cache", ".pyc"}

def texto_normalizado(valor):
    valor = str(valor).lower()
    valor = valor.replace("ç", "c")
    for antigo, novo in (
        ("á", "a"), ("à", "a"), ("â", "a"), ("ã", "a"),
        ("é", "e"), ("ê", "e"), ("í", "i"), ("ó", "o"),
        ("ô", "o"), ("õ", "o"), ("ú", "u"),
    ):
        valor = valor.replace(antigo, novo)
    return valor

def caminho_relativo(caminho):
    try:
        return caminho.relative_to(RAIZ)
    except ValueError:
        return caminho

def nome_pasta_protegida(nome):
    """Identifica pastas protegidas sem diferenciar acentos ou caixa."""
    nome_normalizado = texto_normalizado(nome).replace("_", " ").replace("-", " ")
    nome_normalizado = " ".join(nome_normalizado.split())
    return any(termo in nome_normalizado for termo in TERMOS_PASTAS_PROTEGIDAS)

def caminho_tem_pasta_protegida(caminho):
    """Verifica somente componentes dentro da Área de Trabalho (não o perfil do Windows)."""
    relativo = caminho_relativo(caminho)
    partes = relativo.parts[:-1] if caminho.is_file() else relativo.parts
    return any(nome_pasta_protegida(parte) for parte in partes)

def esta_protegido(caminho):
    """Protege shutdown.bat e todos os itens dentro das pastas protegidas."""
    return (caminho.name.casefold() == ARQUIVO_PROTEGIDO.casefold()
            or caminho_tem_pasta_protegida(caminho))

def pasta_lagora_ou_ancestral(caminho):
    """Mantém o nome histórico da função; agora cobre todas as pastas protegidas."""
    return caminho_tem_pasta_protegida(caminho)

def ano_no_texto(texto):
    encontrados = re.findall(r"(?<!\d)(20\d{2}|19\d{2})(?!\d)", texto)
    # Evita usar anos de versões técnicas, mas mantém anos documentais.
    return encontrados[0] if encontrados else None

def classificar(arquivo):
    """Retorna uma pasta de destino relativa à raiz, sem alterar o nome."""
    nome = texto_normalizado(arquivo.name)
    caminho = texto_normalizado(str(caminho_relativo(arquivo)))
    ext = arquivo.suffix.lower()

    # A ordem importa: termos mais específicos vêm primeiro.
    if any(x in caminho for x in ("censo", "educacenso", "relacaoalunoescola",
                                  "relacaoprofissional", "relacaoturma")):
        categoria = "censo"
    elif any(x in caminho for x in ("frequencia", "presenca", "baixa_freq",
                                    "baixa freq", "faltas", "infrequ")):
        categoria = "frequencia"
    elif any(x in caminho for x in ("ata", "resultado final", "boletim",
                                    "historico escolar", "historicos", "histórico")):
        categoria = "atas"
    elif any(x in caminho for x in ("atestado", "declaracao", "declaração",
                                    "matricula", "matrícula", "escolaridade",
                                    "vaga", "transferencia", "transferência")):
        categoria = "atestados"
    elif any(x in caminho for x in ("inclusao", "inclusão", "monitor", "auxilio",
                                    "auxílio", "adaptacao", "adaptação", "paee")):
        categoria = "inclusao"
    elif any(x in caminho for x in ("passeio", "autorizacao", "autorização",
                                    "proerd", "feira do livro", "uniforme")):
        categoria = "passeio"
    elif any(x in caminho for x in ("prestacao de contas", "prestação de contas",
                                    "pdde", "financeiro", "nota fiscal", "recibo")):
        categoria = "financeiro"
    elif any(x in caminho for x in ("biblioteca", "livros", "retiradas", "devolucoes",
                                    "devoluções", "sistema_biblioteca")):
        categoria = "biblioteca"
    elif any(x in caminho for x in ("aluno", "alunos", "estudante", "estudantes",
                                    "responsavel", "responsável", "crianca", "criança",
                                    "fotos criancas", "fotos crianças", "dossier")):
        categoria = "alunos"
    elif any(x in caminho for x in ("escola", "secretaria", "sargento raymundo",
                                    "funcionario", "funcionário", "professor",
                                    "monitoras", "profs_auxiliares", "caminho escolar")):
        categoria = "escola"
    elif any(x in caminho for x in ("matematica", "matemática", "faculdade",
                                    "trabalho fátima", "trabalho fatima",
                                    "livroingles", "livro ingles", "fisica", "quimica")):
        categoria = "trabalho"
    elif ext in EXTENSOES_TEMPORARIOS or nome.startswith(("~$", "~wrl", ".~")):
        categoria = "temporarios"
    elif ext == ".pdf":
        categoria = "pdf"
    elif ext in {".doc", ".docx", ".odt", ".rtf", ".txt", ".md"}:
        categoria = "word"
    elif ext in {".xls", ".xlsx", ".xlsm", ".ods", ".csv"}:
        categoria = "planilhas"
    elif ext in {".json", ".xml", ".yaml", ".yml", ".toml"}:
        categoria = "dados"
    elif ext in EXTENSOES_INSTALADORES:
        categoria = "instaladores"
    elif ext in EXTENSOES_COMPACTADOS:
        categoria = "compactados"
    elif ext in EXTENSOES_IMAGEM:
        categoria = "imagens"
    elif ext in EXTENSOES_AUDIO_VIDEO:
        categoria = "midia"
    elif ext in EXTENSOES_CODIGO:
        categoria = "programacao"
    elif any(x in nome for x in ("ponto", "autoponto", "planilha_a4", "planilhas_a4")):
        categoria = "escola"
    else:
        categoria = "documentos"

    base = PASTA_DESTINO_GERAL[categoria]

    # Se houver ano no caminho/nome, agrupa por ano para facilitar a busca.
    ano = ano_no_texto(caminho)
    if ano and categoria in {"censo", "frequencia", "atas", "atestados",
                             "inclusao", "passeio", "financeiro", "escola"}:
        return Path(base) / ano
    return Path(base)

def destino_sem_colisao(destino):
    """Nunca sobrescreve: adiciona sufixo numérico quando o nome já existe."""
    if not destino.exists():
        return destino
    pasta, nome, sufixo = destino.parent, destino.stem, destino.suffix
    contador = 2
    while True:
        candidato = pasta / f"{nome} ({contador}){sufixo}"
        if not candidato.exists():
            return candidato
        contador += 1

def descobrir_arquivos():
    arquivos = []
    for atual, diretorios, nomes in os.walk(RAIZ, topdown=True):
        atual_path = Path(atual)
        # Não entra em links simbólicos nem em diretórios protegidos.
        diretorios[:] = sorted(
            d for d in diretorios
            if not (atual_path / d).is_symlink()
            and not nome_pasta_protegida(d)
        )
        for nome in nomes:
            arquivo = atual_path / nome
            if arquivo.is_symlink():
                continue
            if nome.casefold() in {NOME_RELATORIO_INVENTARIO.casefold(),
                                   NOME_LOG.casefold()}:
                continue
            if arquivo.name == NOME_SCRIPT:
                # O script pode estar no próprio Desktop; ele também será
                # classificado, mas não deve mover a si mesmo durante a execução.
                continue
            if esta_protegido(arquivo) or pasta_lagora_ou_ancestral(arquivo):
                continue
            arquivos.append(arquivo)
    return sorted(arquivos, key=lambda p: str(p).casefold())

def pastas_protegidas_ou_com_protegidos():
    """Retorna pastas que precisam permanecer para preservar shutdown.bat."""
    manter = set()
    for atual, diretorios, nomes in os.walk(RAIZ, topdown=True):
        atual_path = Path(atual)
        diretorios[:] = [
            d for d in diretorios
            if not (atual_path / d).is_symlink()
            and not nome_pasta_protegida(d)
        ]
        for nome in nomes:
            if nome.casefold() == ARQUIVO_PROTEGIDO.casefold():
                caminho = atual_path / nome
                pai = caminho.parent
                while pai != RAIZ and RAIZ in pai.parents:
                    manter.add(pai)
                    pai = pai.parent
    # Preserva as pastas lagora e tudo o que estiver dentro delas.
    for atual, diretorios, _ in os.walk(RAIZ, topdown=True):
        atual_path = Path(atual)
        for d in list(diretorios):
            if nome_pasta_protegida(d):
                pasta = atual_path / d
                # Preserva a pasta e todas as subpastas, inclusive as vazias.
                for raiz_protegida, subpastas, _ in os.walk(pasta, topdown=True):
                    raiz_protegida_path = Path(raiz_protegida)
                    manter.add(raiz_protegida_path)
                    subpastas[:] = [
                        sub for sub in subpastas
                        if not (raiz_protegida_path / sub).is_symlink()
                    ]
                pai = pasta.parent
                while pai != RAIZ and RAIZ in pai.parents:
                    manter.add(pai)
                    pai = pai.parent
                diretorios.remove(d)
    return manter

def main():
    print("=" * 68)
    print("ORGANIZADOR DA ÁREA DE TRABALHO")
    print("=" * 68)
    print(f"Área de trabalho alvo: {RAIZ}")
    if Path(__file__).resolve().parent != RAIZ:
        print(f"Observação: o script está em {Path(__file__).resolve().parent}")
    print("Modo inicial: SIMULAÇÃO — nenhum arquivo será movido ainda.")
    print("Proteções: pastas com 'lagora', 'censo' ou 'sistema presença', e todos os shutdown.bat.")
    print()

    arquivos = descobrir_arquivos()
    plano = []
    for origem in arquivos:
        destino_relativo = classificar(origem)
        destino = RAIZ / destino_relativo / origem.name
        destino = destino_sem_colisao(destino)
        if origem.resolve() == destino.resolve():
            continue
        plano.append((origem, destino))

    print(f"Arquivos que podem ser organizados: {len(plano)}")
    print("Exemplos do plano:")
    for origem, destino in plano[:35]:
        print(f"  {caminho_relativo(origem)}")
        print(f"    -> {caminho_relativo(destino)}")
    if len(plano) > 35:
        print(f"  ... e mais {len(plano) - 35} arquivos.")
    print()
    print("As pastas protegidas ('lagora', 'censo', 'sistema presença') e todos os shutdown.bat ficam no lugar.")
    print("Pastas só serão removidas se ficarem vazias.")
    print("Arquivos com nomes repetidos receberão sufixos (2), (3) etc.")
    print()

    resposta = input(
        "Para executar as mudanças, digite ORGANIZAR. "
        "Qualquer outra resposta cancela: "
    ).strip()
    if resposta != "ORGANIZAR":
        print("Cancelado. Nenhum arquivo foi movido.")
        return

    registro = RAIZ / NOME_LOG
    movimentos = 0
    falhas = []

    with registro.open("w", newline="", encoding="utf-8-sig") as csvfile:
        writer = csv.writer(csvfile, delimiter=";")
        writer.writerow(["Data/hora", "Origem", "Destino", "Status", "Detalhe"])
        for numero, (origem, destino_planejado) in enumerate(plano, start=1):
            try:
                # Confere novamente as proteções antes de cada movimento.
                if esta_protegido(origem) or pasta_lagora_ou_ancestral(origem):
                    writer.writerow([datetime.now().isoformat(), str(origem),
                                     "", "IGNORADO", "Caminho protegido"])
                    continue

                destino_planejado.parent.mkdir(parents=True, exist_ok=True)
                destino = destino_sem_colisao(destino_planejado)
                shutil.move(str(origem), str(destino))
                movimentos += 1
                writer.writerow([
                    datetime.now().isoformat(),
                    str(caminho_relativo(origem)),
                    str(caminho_relativo(destino)),
                    "MOVIDO", ""
                ])
                if numero % 50 == 0 or numero == len(plano):
                    print(f"Movidos {numero}/{len(plano)} itens planejados...")
            except Exception as exc:
                falhas.append((str(origem), str(exc)))
                writer.writerow([
                    datetime.now().isoformat(), str(origem),
                    str(destino_planejado), "ERRO", repr(exc)
                ])

        # Registra os itens protegidos encontrados.
        for atual, diretorios, nomes in os.walk(RAIZ, topdown=True):
            atual_path = Path(atual)
            diretorios[:] = [
                d for d in diretorios
                if not (atual_path / d).is_symlink()
                and not nome_pasta_protegida(d)
            ]
            for nome in nomes:
                if nome.casefold() == ARQUIVO_PROTEGIDO.casefold():
                    writer.writerow([
                        datetime.now().isoformat(),
                        str(caminho_relativo(atual_path / nome)),
                        str(caminho_relativo(atual_path / nome)),
                        "PROTEGIDO", "Não mover nem renomear"
                    ])

    # Apaga apenas diretórios vazios, de baixo para cima.
    manter = pastas_protegidas_ou_com_protegidos()
    pastas_removidas = 0
    pastas_com_erro = 0
    for atual, diretorios, _ in os.walk(RAIZ, topdown=False):
        pasta = Path(atual)
        if pasta == RAIZ or pasta in manter:
            continue
        # Nunca remover a pasta que contém o próprio script.
        try:
            pasta.relative_to(Path(__file__).resolve().parent)
            if pasta == RAIZ:
                continue
        except ValueError:
            pass
        try:
            if not any(pasta.iterdir()):
                pasta.rmdir()
                pastas_removidas += 1
        except OSError:
            pastas_com_erro += 1

    print()
    print("=" * 68)
    print("ORGANIZAÇÃO CONCLUÍDA")
    print(f"Arquivos movidos: {movimentos}")
    print(f"Pastas vazias removidas: {pastas_removidas}")
    print(f"Falhas ao mover: {len(falhas)}")
    print(f"Pastas que não puderam ser removidas: {pastas_com_erro}")
    print(f"Registro: {registro}")
    print("=" * 68)

    if falhas:
        print("\nPrimeiras falhas:")
        for origem, erro in falhas[:20]:
            print(f"- {origem}: {erro}")
        print("\nConsulte o CSV para ver o registro completo.")

    input("\nPressione Enter para fechar...")

if __name__ == "__main__":
    main()
