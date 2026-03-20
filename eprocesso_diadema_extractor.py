"""
╔══════════════════════════════════════════════════════════════════════╗
║  EXTRATOR EPROCESSO DIADEMA — v4.0 (produção)                       ║
║  Softplan SPW/CPAV — https://eprocesso.diadema.sp.gov.br            ║
╠══════════════════════════════════════════════════════════════════════╣
║  MAPA DE ENDPOINTS (todos verificados ao vivo):                      ║
║                                                                      ║
║  [AUTH]                                                              ║
║   POST /login/j_security_check                                       ║
║   Cookies obrigatórios: JSESSIONID + SPJSSOID                       ║
║                                                                      ║
║  [FILA DE TRABALHO]                                                  ║
║   GET  /cpavFilaTrabalho/abrirFilaTrabalho.do                        ║
║     → HTML com JSON inline passado a new FilaHandler({...})          ║
║     → charset: windows-1252                                          ║
║     → params: cdUsuarioFila, view, cdVisualizacao=2, tipo           ║
║     → IDs no formato: {tipo}_{ano}_{numero}_{vol}-{seq}             ║
║     → 38 campos por processo no JSON                                 ║
║   POST /cpavFilaTrabalho/consultarFilaTrabalho.do (com filtros)      ║
║                                                                      ║
║  [DETALHE DO PROCESSO — API REST cpa-core-backend]                   ║
║   GET  /cpav/visualizarProcesso.do?chaveProc={base64(key)}           ║
║     → Extrai cdProcesso (UUID) dos hidden inputs                     ║
║   GET  /cpa-core-backend/s/processo/{cdProcesso}/dados-basicos       ║
║     → 56 campos + interessados[] + fluxo{} + consTramiProc{}        ║
║   GET  /cpa-core-backend/s/processo/{cdProcesso}/tramitacoes         ║
║     → array de tramitações (17 campos cada)                          ║
║   GET  /cpa-core-backend/s/processo/consultar/{cdProcesso}           ║
║     → vinculacoes[], juncoes[], prazos[], lembretes[], etc.           ║
║   GET  /cpa-core-backend/s/processo/{cdProcesso}/tipo-processo       ║
║   GET  /cpa-core-backend/processo-fluxo/status/{cdProcesso}          ║
║                                                                      ║
║  [PAINEL DO USUÁRIO — REST]                                          ║
║   POST /cpa-core-backend/processos/ultimosacessos/meus               ║
║     → body: {cdOrgao, quantidade}                                    ║
║     → resposta: {pageNumber, pageSize, totalElements, result:[]}     ║
║   POST /cpa-core-backend/processos/favoritos/meus                    ║
║     → mesma estrutura                                                 ║
║   GET  /ctn-backend/central/notification/user/count/not-displayed    ║
║     → REQUER header Referer: https://...diadema.../portal/           ║
║     → retorna número inteiro simples                                  ║
║   GET  /portal-web-backend/api/v1/usuario                            ║
║   GET  /portal-web-backend/api/v1/menus                              ║
║   GET  /portal-web-backend/api/v1/dashboard/charts                   ║
╚══════════════════════════════════════════════════════════════════════╝
Requisitos:
    pip install requests beautifulsoup4 pandas openpyxl lxml
"""
import requests
from bs4 import BeautifulSoup
import pandas as pd
import json
import os
import re
import base64
import time
from datetime import datetime
from typing import Optional

# ──────────────────────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ──────────────────────────────────────────────────────────────────────
BASE_URL   = "https://eprocesso.diadema.sp.gov.br"
LOGIN_URL  = f"{BASE_URL}/login/j_security_check"
PORTAL_URL = f"{BASE_URL}/portal/"
FILA_ENCODING = "windows-1252"
CD_ORGAO      = "PMDIADEMA"

HEADERS_PADRAO = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
    "Origin":          BASE_URL,
    "Referer":         f"{BASE_URL}/portal/",   # obrigatório p/ ctn-backend
}


# ──────────────────────────────────────────────────────────────────────
# EXTRATOR PRINCIPAL
# ──────────────────────────────────────────────────────────────────────
class EprocessoExtractor:
    def __init__(self, usuario: str, senha: str):
        self.usuario = usuario
        self.senha   = senha
        self.session = requests.Session()
        self.session.headers.update(HEADERS_PADRAO)
        self.logado  = False

    # ──────────────────────────────────────────────────────────────────
    # LOGIN
    # ──────────────────────────────────────────────────────────────────
    def login(self) -> bool:
        """
        Autentica via j_security_check.
        Após login bem-sucedido a sessão contém JSESSIONID + SPJSSOID.
        """
        print(f"[INFO] Fazendo login como: {self.usuario}")
        resp = self.session.get(f"{BASE_URL}/login/", timeout=30)
        resp.raise_for_status()

        payload = {
            "j_username":    self.usuario,
            "j_password":    self.senha,
            "senhaExpirada": "",
            "cdConfigLDAP":  "",
        }
        resp = self.session.post(
            LOGIN_URL, data=payload, allow_redirects=True, timeout=30
        )
        resp.raise_for_status()

        if "/portal/" in resp.url or "/cpav/" in resp.url:
            print(f"[OK] Login bem-sucedido → {resp.url}")
            self.logado = True
            return True

        r2 = self.session.get(PORTAL_URL, timeout=30, allow_redirects=True)
        if "/login/" not in r2.url:
            print(f"[OK] Sessão ativa → {r2.url}")
            self.logado = True
            return True

        print("[ERRO] Login falhou.")
        return False

    def _check_login(self):
        if not self.logado:
            raise RuntimeError("Execute login() primeiro.")

    def _get(self, url: str, params: dict = None,
             encoding: str = None) -> requests.Response:
        resp = self.session.get(url, params=params, timeout=30)
        resp.raise_for_status()
        if "/login/" in resp.url:
            print("[AVISO] Sessão expirada — reautenticando...")
            self.logado = False
            self.login()
            resp = self.session.get(url, params=params, timeout=30)
            resp.raise_for_status()
        if encoding:
            resp.encoding = encoding
        return resp

    def _post(self, url: str,
              data: dict = None,
              json_data: dict = None) -> requests.Response:
        resp = self.session.post(url, data=data, json=json_data, timeout=30)
        resp.raise_for_status()
        return resp

    # ──────────────────────────────────────────────────────────────────
    # FILA DE TRABALHO
    # ──────────────────────────────────────────────────────────────────
    def fila_de_trabalho(
        self,
        view: str    = "Todos",
        tipo: str    = "Digitais",
        usuario: str = None,
    ) -> list[dict]:
        """
        Retorna todos os processos da Fila de Trabalho.
        O servidor entrega o HTML com slots vazios (empty-0…N) e um
        script inline com o JSON de todos os processos:
            var filaHandler = new FilaHandler({'1_2026_4370_1-1': {
                "nuProcesso": 4370, "nuAno": "2026",
                "numeroFormatado": "PMDI 00004370/2026",
                ...38 campos...
            }, ...})
        O JS clona o template e preenche os slots no browser.
        O requests recebe o HTML completo com o JSON — basta extraí-lo.
        """
        self._check_login()
        if usuario is None:
            usuario = self.usuario.upper()

        url = f"{BASE_URL}/cpavFilaTrabalho/abrirFilaTrabalho.do"
        params = {
            "cdUsuarioFila":  usuario,
            "view":           view,
            "taskName":       "",
            "cdVisualizacao": "2",
            "tipo":           tipo,
        }
        print(f"[INFO] Fila de Trabalho — view={view}, tipo={tipo}...")
        resp = self._get(url, params=params, encoding=FILA_ENCODING)
        debug_path = f"fila_debug_{view}_{tipo}.html"
        processos = self._extrair_json_fila(resp.text, debug_html_path=debug_path)
        print(f"[OK] {len(processos)} processos.")
        return processos

    @staticmethod
    def _extrair_json_fila(html: str, debug_html_path: str = None) -> list[dict]:
        """
        Extrai o JSON inline passado para new FilaHandler({...}).
        Se debug_html_path for informado, salva o HTML bruto para análise.
        """
        if debug_html_path:
            with open(debug_html_path, "w", encoding="utf-8") as f:
                f.write(html)
            print(f"[DEBUG] HTML bruto salvo em: {debug_html_path}")

        # Contagem prévia de IDs para diagnóstico
        ids_no_html = re.findall(r"'(\d+_\d+_\d+_\d+-\d+)'", html)
        if ids_no_html:
            print(f"[DEBUG] IDs encontrados no HTML bruto: {len(ids_no_html)}")

        marker    = "new FilaHandler("
        idx_start = html.find(marker)
        if idx_start == -1:
            print("[AVISO] Marcador 'new FilaHandler(' não encontrado. Fallback HTML.")
            return EprocessoExtractor._parse_fila_fallback(html)

        obj_start = html.find("{", idx_start + len(marker))
        if obj_start == -1:
            return []

        # Balancear chaves para encontrar o fim do objeto
        depth, in_str, esc, quote_char, obj_end = 0, False, False, None, -1
        for i in range(obj_start, len(html)):
            c = html[i]
            if esc:
                esc = False
                continue
            if c == "\\" and in_str:
                esc = True
                continue
            if in_str:
                if c == quote_char:
                    in_str = False
            else:
                if c in ('"', "'"):
                    in_str, quote_char = True, c
                elif c == "{":
                    depth += 1
                elif c == "}":
                    depth -= 1
                    if depth == 0:
                        obj_end = i
                        break

        if obj_end == -1:
            print("[AVISO] Balanceamento de chaves falhou — objeto truncado. Fallback HTML.")
            return EprocessoExtractor._parse_fila_fallback(html)

        js_obj = html[obj_start : obj_end + 1]

        # Normalizar todas as chaves single-quoted (não apenas IDs de processo)
        js_obj_json = re.sub(
            r"'([^'\\]*(?:\\.[^'\\]*)*)'(\s*:)",
            lambda m: '"' + m.group(1).replace('"', '\\"') + '"' + m.group(2),
            js_obj
        )

        try:
            data = json.loads(js_obj_json)
        except json.JSONDecodeError as e:
            print(f"[AVISO] json.loads falhou ({e}). Tentando ast.literal_eval...")
            try:
                import ast
                data = ast.literal_eval(js_obj)
            except Exception as e2:
                print(f"[ERRO] ast.literal_eval falhou ({e2}). Fallback HTML.")
                return EprocessoExtractor._parse_fila_fallback(html)

        processos = []
        for id_raw, campos in data.items():
            row = {"id_raw": id_raw}
            row.update(campos)
            processos.append(row)

        ids_unicos_html = set(ids_no_html)
        ids_parseados   = {p["id_raw"] for p in processos}
        faltando = ids_unicos_html - ids_parseados
        if faltando:
            print(
                f"[AVISO] Divergência: {len(ids_unicos_html)} IDs únicos no HTML, "
                f"{len(processos)} parseados. Faltando: {faltando}"
            )

        return processos

    @staticmethod
    def _parse_fila_fallback(html: str) -> list[dict]:
        """
        Fallback via BeautifulSoup.
        Ignora os templates vazios (id começa com 'empty').
        """
        soup = BeautifulSoup(html, "lxml")
        col_map = {
            0: "numero",          1: "unidade_enc",
            2: "unidade_atual",   3: "usuario_recebimento",
            4: "encaminhamento",  5: "prazo_encaminhamento",
            6: "detalhamento",    7: "status",
        }
        processos = []
        for item in soup.select(".item-fila"):
            id_raw = item.get("id", "")
            if id_raw.startswith("empty"):
                continue
            row = {
                "id_raw":       id_raw,
                "nao_recebido": "processoNaoRecebido" in item.get("class", []),
            }
            for idx, col in enumerate(item.select(".fila-column")):
                info_div = col.select_one(".info")
                row[col_map.get(idx, f"col_{idx}")] = (
                    info_div.get_text(strip=True) if info_div else ""
                )
            processos.append(row)
        return processos

    def fila_consultar(
        self,
        view: str    = "Todos",
        tipo: str    = "Digitais",
        usuario: str = None,
        filtro: str  = "",
        dt_encam_de:  str = "",
        dt_encam_ate: str = "",
        prazo_de:     str = "",
        prazo_ate:    str = "",
    ) -> list[dict]:
        """Consulta a fila com filtros via POST."""
        self._check_login()
        if usuario is None:
            usuario = self.usuario.upper()

        url  = f"{BASE_URL}/cpavFilaTrabalho/consultarFilaTrabalho.do"
        data = {
            "cdAssunto": "", "cdCategoria": "", "cdVisualizacao": "2",
            "cdGrupoassunto": "", "cdSetorAtual": "", "cdUsuarioFila": usuario,
            "nmTarefa": "", "tipo": tipo, "view": view,
            "filtro.cdAssunto": "", "filtro.cdVisualizacao": "2",
            "filtro.cdGrupoassunto": "", "filtro.cdUsuario": usuario,
            "filtro.nmTarefa": "", "filtro.tipo": tipo, "filtro.view": view,
            "filtro.texto": filtro,
            "filtro.dtEncaminhamentoDe":  dt_encam_de,
            "filtro.dtEncaminhamentoAte": dt_encam_ate,
            "filtro.prazoDe":  prazo_de,
            "filtro.prazoAte": prazo_ate,
        }
        resp = self._post(url, data=data)
        resp.encoding = FILA_ENCODING
        return self._extrair_json_fila(resp.text, debug_html_path="fila_debug_consultar.html")

    # ──────────────────────────────────────────────────────────────────
    # DETALHE COMPLETO DO PROCESSO
    # ──────────────────────────────────────────────────────────────────
    @staticmethod
    def _chave_para_b64(chave: str) -> str:
        """
        Converte a chave do processo para base64 (sem padding).
        Ex: '1_2026_4370_1-1'  →  'MV8yMDI2XzQzNzBfMS0x'
        """
        return base64.b64encode(chave.encode()).decode().rstrip("=")

    def _extrair_cd_processo(self, chave: str) -> Optional[str]:
        """
        Abre a página de visualização e extrai o cdProcesso (UUID).
        """
        b64  = self._chave_para_b64(chave)
        url  = f"{BASE_URL}/cpav/visualizarProcesso.do"
        resp = self._get(url, params={"chaveProc": b64})
        soup = BeautifulSoup(resp.text, "lxml")
        el   = soup.find("input", {"name": "cdProcesso"})
        if el and el.get("value"):
            return el["value"]
        match = re.search(
            r'name="cdProcesso"\s+value="([0-9a-f-]{36})"', resp.text
        )
        return match.group(1) if match else None

    def detalhe_processo(
        self,
        numero: int,
        ano: int,
        volume: int = 1,
    ) -> dict:
        """
        Retorna o detalhe completo de um processo via API REST.
        """
        self._check_login()
        chave       = f"1_{ano}_{numero}_{volume}-1"
        cd_processo = self._extrair_cd_processo(chave)
        if not cd_processo:
            raise ValueError(f"Processo {numero}/{ano} não encontrado.")

        print(f"[INFO] Processo {numero}/{ano} → cdProcesso={cd_processo}")

        db = self._get(
            f"{BASE_URL}/cpa-core-backend/s/processo/{cd_processo}/dados-basicos"
        ).json()

        tramitacoes = self._get(
            f"{BASE_URL}/cpa-core-backend/s/processo/{cd_processo}/tramitacoes"
        ).json()

        consultar = self._get(
            f"{BASE_URL}/cpa-core-backend/s/processo/consultar/{cd_processo}"
        ).json()

        tipo = self._get(
            f"{BASE_URL}/cpa-core-backend/s/processo/{cd_processo}/tipo-processo"
        ).json()

        try:
            fluxo = self._get(
                f"{BASE_URL}/cpa-core-backend/processo-fluxo/status/{cd_processo}"
            ).json()
        except Exception:
            fluxo = {}

        return {
            "cd_processo":             cd_processo,
            "chave":                   chave,
            "dados_basicos":           db,
            "tramitacoes":             tramitacoes,
            "consultar":               consultar,
            "tipo_processo":           tipo,
            "fluxo_status":            fluxo,
            "numero_formatado":        db.get("numeroFormatado"),
            "orgao":                   db.get("siglaOrgaoSetorProcesso"),
            "situacao":                db.get("nomeSituacao"),
            "setor_atual":             db.get("siglaOrgaoSetorAtual"),
            "setor_atual_nome":        db.get("nomeOrgaoSetorAtual"),
            "data_entrada":            db.get("dataEntrada"),
            "hora_entrada":            db.get("horaEntrada"),
            "data_autuacao":           db.get("dataAutuacao"),
            "classificacao":           db.get("nomeClasse"),
            "codigo_classificacao":    db.get("codigoCompletoClassificacao"),
            "detalhamento":            db.get("detalhamento"),
            "interessado_principal":   db.get("nomeInteressadoPrincipal"),
            "cpf_cnpj_interessado":    db.get("identificadorInteressadoPrincipal"),
            "procedencia":             db.get("procedencia"),
            "origem_cadastro":         db.get("origemCadastro", db.get("tpOrigemCadastro")),
            "tipo_processo_nome":      db.get("nomeTipoProcesso"),
            "controle_acesso":         db.get("nomeModeloControleAcesso"),
            "usuario_cadastro":        db.get("usuarioCadastro"),
            "usuario_autuacao":        db.get("usuarioAutuacao"),
            "interessados":            db.get("interessados", []),
            "tramitacao_atual":        db.get("consTramiProc", {}),
            "vinculacoes":             consultar.get("vinculacoes", []),
            "juncoes":                 consultar.get("juncoes", []),
            "prazos":                  consultar.get("prazos", []),
            "lembretes":               consultar.get("lembretes", []),
            "arquivamentos":           consultar.get("arquivamentos", []),
            "cancelamentos":           consultar.get("cancelamentos", []),
            "processos_externos":      consultar.get("processosExternos", []),
        }

    def detalhe_em_lote(
        self,
        processos_fila: list[dict],
    ) -> list[dict]:
        """
        Extrai detalhes completos para uma lista de processos da fila.
        """
        resultados = []
        for p in processos_fila:
            id_raw = p.get("id_raw", "")
            partes = id_raw.replace("-", "_").split("_")
            if len(partes) < 4:
                continue
            try:
                numero = int(partes[2])
                ano    = int(partes[1])
                volume = int(partes[3])
            except (ValueError, IndexError):
                continue
            try:
                detalhe = self.detalhe_processo(numero, ano, volume)
                detalhe["_fila"] = p
                resultados.append(detalhe)
                print(f"  [OK] {detalhe.get('numero_formatado')}")
            except Exception as e:
                print(f"  [ERRO] {id_raw}: {e}")
            time.sleep(0.3)
        return resultados

    # ──────────────────────────────────────────────────────────────────
    # ÚLTIMOS ACESSOS E FAVORITOS
    # ──────────────────────────────────────────────────────────────────
    def ultimos_acessos(self, quantidade: int = 50) -> list[dict]:
        """
        POST /cpa-core-backend/processos/ultimosacessos/meus
        Resposta real:
        {
          "pageNumber": 0, "pageSize": 10, "totalElements": 10,
          "result": [
            {"cdUsuario": "...", "cdProcesso": "uuid",
             "nmProcesso": "PMDI 00035923/2025", "descricao": "...",
             "interessados": "...", "classificacao": "...",
             "processoPK": "...", "dtCriacao": "..."},
            ...
          ]
        }
        """
        self._check_login()
        url  = f"{BASE_URL}/cpa-core-backend/processos/ultimosacessos/meus"
        body = {"cdOrgao": CD_ORGAO, "quantidade": quantidade}
        print(f"[INFO] Últimos {quantidade} acessos...")
        resp = self._post(url, json_data=body)
        return self._extrair_result(resp.json())

    def favoritos(self, quantidade: int = 50) -> list[dict]:
        """
        POST /cpa-core-backend/processos/favoritos/meus
        Mesma estrutura que ultimos_acessos().
        """
        self._check_login()
        url  = f"{BASE_URL}/cpa-core-backend/processos/favoritos/meus"
        body = {"cdOrgao": CD_ORGAO, "quantidade": quantidade}
        print("[INFO] Favoritos...")
        resp = self._post(url, json_data=body)
        return self._extrair_result(resp.json())

    @staticmethod
    def _extrair_result(payload) -> list[dict]:
        """
        Extrai a lista de processos do envelope de paginação.
        Suporta dict com "result" (formato real) ou lista na raiz.
        """
        if isinstance(payload, list):
            return payload
        if isinstance(payload, dict):
            result = payload.get("result", [])
            if isinstance(result, list):
                return result
            if isinstance(result, dict):
                return [result]
        return []

    # ──────────────────────────────────────────────────────────────────
    # NOTIFICAÇÕES
    # ──────────────────────────────────────────────────────────────────
    def contar_notificacoes(self) -> int:
        """
        GET /ctn-backend/central/notification/user/count/not-displayed
        Exige header Referer apontando para o portal (já configurado).
        Retorna número inteiro (ex: 107).
        """
        self._check_login()
        url  = f"{BASE_URL}/ctn-backend/central/notification/user/count/not-displayed"
        resp = self._get(url)
        try:
            return int(resp.text.strip())
        except ValueError:
            data = resp.json()
            return data.get("count", -1) if isinstance(data, dict) else -1

    # ──────────────────────────────────────────────────────────────────
    # PORTAL — DADOS DO USUÁRIO E MENUS
    # ──────────────────────────────────────────────────────────────────
    def dados_usuario(self) -> dict:
        """
        GET /portal-web-backend/api/v1/usuario
        Retorna: nome, login, deEmail, dtAlteracaoSenha,
                 dtUltimoLoginSucesso, dtUltimoLoginFalho,
                 role, isLdap, isLoginOpenId
        """
        self._check_login()
        return self._get(
            f"{BASE_URL}/portal-web-backend/api/v1/usuario"
        ).json()

    def menus(self) -> list[dict]:
        """
        GET /portal-web-backend/api/v1/menus
        Retorna estrutura completa de menus do portal.
        """
        self._check_login()
        return self._get(
            f"{BASE_URL}/portal-web-backend/api/v1/menus"
        ).json()

    # ──────────────────────────────────────────────────────────────────
    # EXPORTAÇÃO
    # ──────────────────────────────────────────────────────────────────
    @staticmethod
    def para_csv(dados: list[dict], arquivo: str = "resultado.csv"):
        if not dados:
            print("[AVISO] Nenhum dado para exportar.")
            return
        pd.DataFrame(dados).to_csv(arquivo, index=False, encoding="utf-8-sig")
        print(f"[OK] CSV → {arquivo} ({len(dados)} registros)")

    @staticmethod
    def para_excel(abas: dict[str, list], arquivo: str = "resultado.xlsx"):
        """abas = {"Nome": [lista_de_dicts], ...}"""
        with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
            for nome, dados in abas.items():
                if dados:
                    pd.json_normalize(dados).to_excel(
                        writer, sheet_name=nome[:31], index=False
                    )
                    print(f"    aba '{nome}': {len(dados)} registros")
                else:
                    print(f"    aba '{nome}': sem dados")
        print(f"[OK] Excel → {arquivo}")

    @staticmethod
    def para_json(dados, arquivo: str = "resultado.json"):
        with open(arquivo, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2, default=str)
        print(f"[OK] JSON → {arquivo}")

    @staticmethod
    def tramitacoes_para_df(detalhe: dict) -> pd.DataFrame:
        """Converte a lista de tramitações de um detalhe em DataFrame."""
        trams = detalhe.get("tramitacoes", [])
        if not trams:
            return pd.DataFrame()
        return pd.json_normalize(trams)


# ──────────────────────────────────────────────────────────────────────
# USO
# ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    USUARIO = os.getenv("EPROCESSO_USER", "gabriel.gomes")
    SENHA   = os.getenv("EPROCESSO_PASS", "Sre2710*")

    ext = EprocessoExtractor(USUARIO, SENHA)
    if not ext.login():
        exit(1)

    # ── Fila de Trabalho ──────────────────────────────────────────────
    fila_todos     = ext.fila_de_trabalho(view="Todos",         tipo="Digitais")
    fila_meus      = ext.fila_de_trabalho(view="MeusProcessos", tipo="Digitais")
    fila_recebidos = ext.fila_de_trabalho(view="Recebidos",     tipo="Digitais")

    ext.para_csv(fila_todos, "fila_todos.csv")

    # ── Últimos acessos e favoritos ───────────────────────────────────
    ultimos   = ext.ultimos_acessos(quantidade=50)
    favoritos = ext.favoritos()

    # ── Notificações ──────────────────────────────────────────────────
    n = ext.contar_notificacoes()
    print(f"[INFO] Notificações não lidas: {n}")

    # ── Dados do usuário ──────────────────────────────────────────────
    try:
        usuario_info = ext.dados_usuario()
        print(f"[INFO] Usuário: {usuario_info.get('nome')} | {usuario_info.get('deEmail')}")
    except Exception as e:
        print(f"[AVISO] dados_usuario: {e}")
        usuario_info = {}

    # ── Exportação consolidada ────────────────────────────────────────
    hoje = datetime.today().strftime("%Y%m%d")
    ext.para_excel(
        {
            "Fila - Todos":     fila_todos,
            "Fila - Meus":      fila_meus,
            "Fila - Recebidos": fila_recebidos,
            "Ultimos Acessos":  ultimos,
            "Favoritos":        favoritos,
        },
        arquivo=f"eprocesso_diadema_{hoje}.xlsx",
    )

    # ── Exemplo: detalhe de um processo específico ────────────────────
    # detalhe = ext.detalhe_processo(numero=4370, ano=2026)
    # ext.para_json(detalhe, "detalhe_4370_2026.json")
    # print(pd.DataFrame([detalhe]).T)

    # ── Exemplo: detalhes em lote de toda a fila ─────────────────────
    # detalhes = ext.detalhe_em_lote(fila_todos)
    # ext.para_excel({"Detalhes Fila": detalhes}, "fila_detalhada.xlsx")
