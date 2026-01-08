"""
Bot de Circularização via WhatsApp
Versão: 1.1 - Revisado e robustecido
Desenvolvido para Auditoria Interna
"""

import os
import time
import logging
from datetime import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

# -------------------------
# Configuração de logging
# -------------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('circularizacao_bot.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# -------------------------
# Classe do Bot
# -------------------------
class CircularizacaoBot:
    def __init__(
        self,
        planilha_clientes='clientes_circularizacao.xlsx',
        profile_path=None,
        chrome_driver_path=None,
        max_wait_login=120
    ):
        self.planilha_clientes = planilha_clientes
        self.driver = None
        self.respostas = []
        self.chrome_driver_path = chrome_driver_path
        self.max_wait_login = max_wait_login

        # Diretório para armazenar sessão do WhatsApp
        if profile_path is None:
            self.profile_path = os.path.join(os.getcwd(), 'whatsapp_session')
        else:
            self.profile_path = profile_path
        os.makedirs(self.profile_path, exist_ok=True)

    # -------------------------
    # Utilitários
    # -------------------------
    def _safe_find(self, by, selector, timeout=10):
        """Tenta encontrar elemento e retorna elemento ou None."""
        try:
            return WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((by, selector))
            )
        except Exception:
            return None

    def _safe_find_all(self, by, selector, timeout=5):
        """Retorna lista de elementos (pode ser vazia)."""
        try:
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((by, selector))
            )
        except Exception:
            pass
        try:
            return self.driver.find_elements(by, selector)
        except Exception:
            return []

    def _sanitize_phone(self, raw):
        """
        Normaliza telefone:
         - remove espaços, traços, parênteses
         - se não começar com 55, assume DDI BR e adiciona
        """
        if pd.isna(raw):
            return None
        s = str(raw)
        # remove caracteres não numéricos
        digits = ''.join(ch for ch in s if ch.isdigit())
        if not digits:
            return None
        # se tiver 8-9 dígitos (sem DDD) -> não é suficiente; retornamos como está
        # Garantir DDI 55 para Brasil se for menor que 12 dígitos
        if not digits.startswith('55'):
            # evitar adicionar 55 em caso de números com 13+ dígitos (prováveis)
            if len(digits) <= 11:
                digits = '55' + digits
        return digits

    def _sanitize_message(self, text):
        """
        Remove caracteres fora do BMP (ord > 0xFFFF) que causam erro no ChromeDriver.
        Também remove caracteres de controle indesejados (exceto \n e \t).
        Retorna a string sanitizada.
        """
        if text is None:
            return ""
        # Constrói nova string apenas com caracteres ord <= 0xFFFF
        sanitized_chars = []
        removed_count = 0
        for c in str(text):
            oc = ord(c)
            # manter linha e tab, mas remover outros control chars (<=31) exceto \n (10) e \t (9) e \r (13)
            if oc > 0xFFFF:
                removed_count += 1
                continue
            if oc < 32 and oc not in (9, 10, 13):
                # ignora controles não imprimíveis
                removed_count += 1
                continue
            sanitized_chars.append(c)
        sanitized = ''.join(sanitized_chars)
        if removed_count > 0:
            logging.info(f"Sanitização: {removed_count} caracteres removidos da mensagem (possíveis emojis ou controles não-BMP).")
        return sanitized

    # -------------------------
    # Inicialização do WhatsApp
    # -------------------------
    def iniciar_whatsapp(self, headless=False):
        logging.info("Iniciando WhatsApp Web...")
        options = Options()
        # perfil para manter sessão
        options.add_argument(f"--user-data-dir={self.profile_path}")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        if headless:
            # headless não recomendado para leitura de QR, mantemos opção
            options.add_argument("--headless=new")
            options.add_argument("--window-size=1920,1080")
        else:
            options.add_argument("--start-maximized")

        try:
            if self.chrome_driver_path:
                service = Service(self.chrome_driver_path)
                self.driver = webdriver.Chrome(service=service, options=options)
            else:
                self.driver = webdriver.Chrome(options=options)
        except WebDriverException as e:
            logging.error(f"Erro ao iniciar o Chrome WebDriver: {e}")
            raise

        self.driver.get("https://web.whatsapp.com")
        print("\n" + "="*60)
        print("IMPORTANTE: Escaneie o QR Code no WhatsApp (se necessário).")
        print("Aguardando login...")
        print("="*60 + "\n")

        # Esperar por presença do campo de busca (usado como indicador de login)
        login_ok = False
        start = time.time()
        while time.time() - start < self.max_wait_login:
            # tentativa 1: campo de busca por data-tab (varia por versão)
            elem = self._safe_find(By.XPATH, '//div[@contenteditable="true" and (@data-tab="3" or @data-tab="1")]', timeout=5)
            if elem:
                login_ok = True
                break
            # tentativa 2: verificar se a barra lateral está presente
            elem2 = self._safe_find(By.XPATH, '//div[contains(@aria-label,"Search or start new chat") or contains(@class,"app")]',
                                     timeout=5)
            if elem2:
                login_ok = True
                break
            time.sleep(1)

        if not login_ok:
            logging.error("Timeout ao aguardar login no WhatsApp Web.")
            raise TimeoutException("Login WhatsApp Web não detectado dentro do tempo limite.")
        logging.info("Login realizado com sucesso!")
        # aguardar estabilidade do DOM
        time.sleep(2)

    # -------------------------
    # Carregar clientes
    # -------------------------
    def carregar_clientes(self):
        logging.info(f"Carregando planilha: {self.planilha_clientes}")
        if not os.path.exists(self.planilha_clientes):
            logging.error(f"Arquivo não encontrado: {self.planilha_clientes}")
            raise FileNotFoundError(self.planilha_clientes)

        # Suportar Excel e CSV
        try:
            if self.planilha_clientes.lower().endswith(('.xls', '.xlsx')):
                df = pd.read_excel(self.planilha_clientes, dtype=str)
            elif self.planilha_clientes.lower().endswith('.csv'):
                df = pd.read_csv(self.planilha_clientes, dtype=str)
            else:
                df = pd.read_excel(self.planilha_clientes, dtype=str)
        except Exception as e:
            logging.error(f"Erro ao ler a planilha: {e}")
            raise

        # Normalizar nomes de colunas esperadas (tenta mapear por variações simples)
        cols = {c.lower().strip(): c for c in df.columns}
        expected = {}
        # possíveis aliases
        aliases = {
            'nome_cliente': ['nome', 'cliente', 'nome_cliente', 'nome do cliente'],
            'telefone': ['telefone', 'tel', 'celular', 'telefone_cliente', 'telefone do cliente'],
            'notas_fiscais': ['notas_fiscais', 'nf', 'notas', 'notas fiscais', 'notas_fiscais_cliente']
        }
        for key, tries in aliases.items():
            expected[key] = None
            for t in tries:
                if t in cols:
                    expected[key] = cols[t]
                    break

        # verificar que temos pelo menos telefone e nome
        if expected['telefone'] is None or expected['nome_cliente'] is None:
            logging.error("Planilha precisa conter, no mínimo, colunas 'nome_cliente' e 'telefone'."
                          f" Colunas encontradas: {list(df.columns)}")
            raise ValueError("Colunas obrigatórias ausentes na planilha.")

        # renomear para colunas padrão para facilitar uso posterior
        rename_map = {
            expected['nome_cliente']: 'nome_cliente',
            expected['telefone']: 'telefone'
        }
        if expected.get('notas_fiscais'):
            rename_map[expected['notas_fiscais']] = 'notas_fiscais'

        df = df.rename(columns=rename_map)

        # garantir colunas notas_fiscais presentes
        if 'notas_fiscais' not in df.columns:
            df['notas_fiscais'] = ''

        # remover linhas sem telefone
        df['telefone'] = df['telefone'].astype(str).fillna('').replace('nan', '')
        df = df[df['telefone'].str.strip() != '']
        df = df.reset_index(drop=True)

        logging.info(f"{len(df)} clientes carregados")
        return df

    # -------------------------
    # Buscar contato
    # -------------------------
    def buscar_contato(self, numero_telefone, timeout=8):
        """
        Busca um contato pelo número usando a barra de busca do WhatsApp Web.
        Retorna True se o contato foi aberto.
        """
        if not numero_telefone:
            return False

        try:
            # tentar abrir via busca interna
            search_box = None
            possible_search_sel = [
                ('xpath', '//div[@contenteditable="true" and (@data-tab="3" or @data-tab="1")]'),
                ('css', 'div[contenteditable="true"][data-tab]'),
                ('xpath', '//div[@role="textbox" and @contenteditable="true"]')
            ]
            for kind, sel in possible_search_sel:
                if kind == 'xpath':
                    search_box = self._safe_find(By.XPATH, sel, timeout=2)
                else:
                    search_box = self._safe_find(By.CSS_SELECTOR, sel, timeout=2)
                if search_box:
                    break

            if search_box:
                try:
                    search_box.click()
                    time.sleep(0.3)
                except Exception:
                    pass
                # limpar e digitar
                try:
                    search_box.send_keys(Keys.CONTROL + 'a')
                    search_box.send_keys(Keys.BACKSPACE)
                except Exception:
                    try:
                        search_box.clear()
                    except Exception:
                        pass
                time.sleep(0.3)
                search_box.send_keys(numero_telefone)
                time.sleep(2)

                # tentar clicar no contato listado
                try:
                    # tenta título exato
                    contato = WebDriverWait(self.driver, timeout=3).until(
                        EC.element_to_be_clickable((By.XPATH, f'//span[@title="{numero_telefone}"]'))
                    )
                    contato.click()
                    time.sleep(1.2)
                    return True
                except Exception:
                    # tentar any item que contenha o número no título
                    try:
                        items = self.driver.find_elements(By.XPATH, f'//div[contains(@title, "{numero_telefone}")]')
                        if items:
                            items[0].click()
                            time.sleep(1.2)
                            return True
                    except Exception:
                        pass

            # Se falhar pela busca interna, usar fallback via URL de envio (abre o chat diretamente)
            logging.info("Busca interna falhou — tentando abrir chat via URL (/send?phone=)...")
            self.driver.get(f"https://web.whatsapp.com/send?phone={numero_telefone}&text&app_absent=0")
            # aguardar até que a caixa de mensagem apareça (indicando chat aberto)
            started = time.time()
            while time.time() - started < timeout:
                # procurar caixa de mensagem
                msg_box = self._safe_find(By.XPATH, '//div[@contenteditable="true" and (@data-tab="10" or @data-tab="6")]',
                                           timeout=2)
                if msg_box:
                    time.sleep(0.8)
                    return True
                time.sleep(0.5)

            logging.warning(f"Contato não encontrado via busca nem via URL: {numero_telefone}")
            return False

        except Exception as e:
            logging.error(f"Erro ao buscar contato {numero_telefone}: {e}", exc_info=True)
            return False

    # -------------------------
    # Enviar mensagem
    # -------------------------
    def _click_send_button(self):
        """Procura botão de enviar e clica; retorna True se clicou."""
        send_selectors = [
            (By.CSS_SELECTOR, 'button[data-testid="compose-btn-send"]'),
            (By.CSS_SELECTOR, 'span[data-icon="send"]'),
            (By.XPATH, '//button[contains(@aria-label,"Enviar") or contains(@aria-label,"Send")]'),
            (By.CSS_SELECTOR, 'button[title="Enviar"]')
        ]
        for by, sel in send_selectors:
            try:
                btn = self._safe_find(by, sel, timeout=1)
                if btn:
                    try:
                        btn.click()
                    except Exception:
                        # tentar via JS click
                        try:
                            self.driver.execute_script("arguments[0].click();", btn)
                        except Exception:
                            pass
                    return True
            except Exception:
                continue
        return False

    def enviar_mensagem(self, mensagem, timeout=10):
        """
        Envia a mensagem para o chat atualmente aberto.
        ALTERAÇÃO: agora sanitizamos a mensagem para remover caracteres fora do BMP.
        """
        if not mensagem:
            return False
        try:
            # sanitize message to avoid ChromeDriver BMP limitation
            mensagem = self._sanitize_message(mensagem)

            # localizar caixa de mensagem (vários seletores)
            message_box = None
            selectors = [
                ('xpath', '//div[@contenteditable="true" and (@data-tab="10" or @data-tab="6")]'),
                ('xpath', '//div[@role="textbox" and @contenteditable="true"]'),
                ('css', 'div[contenteditable="true"][data-tab]'),
            ]
            for kind, sel in selectors:
                if kind == 'xpath':
                    message_box = self._safe_find(By.XPATH, sel, timeout=2)
                else:
                    message_box = self._safe_find(By.CSS_SELECTOR, sel, timeout=2)
                if message_box:
                    break

            if not message_box:
                logging.error("Caixa de mensagem não encontrada.")
                return False

            # garantir foco
            try:
                message_box.click()
                time.sleep(0.2)
            except Exception:
                pass

            # escrever respeitando quebras de linha
            linhas = str(mensagem).split('\n')
            for i, linha in enumerate(linhas):
                # evitar enviar com enter ao digitar múltiplas linhas
                if linha:
                    message_box.send_keys(linha)
                # se linha vazia, apenas cria a quebra
                if i < len(linhas) - 1:
                    message_box.send_keys(Keys.SHIFT + Keys.ENTER)
                    time.sleep(0.05)

            time.sleep(0.3)

            # Primeiro: tentar clicar no botão de enviar (mais confiável)
            clicked = self._click_send_button()
            if clicked:
                time.sleep(0.6)
                logging.info("Mensagem enviada (clicando no botão enviar).")
                return True

            # Fallback: enviar com Enter
            try:
                message_box.send_keys(Keys.ENTER)
                time.sleep(0.6)
                logging.info("Mensagem enviada (ENTER).")
                return True
            except Exception:
                logging.error("Falha ao enviar mensagem com ENTER e botão enviar indisponível.")
                return False

        except Exception as e:
            logging.error(f"Erro ao enviar mensagem: {e}", exc_info=True)
            return False

    # -------------------------
    # Aguardar resposta
    # -------------------------
    def aguardar_resposta(self, timeout=300, poll_interval=2):
        """
        Aguarda pela próxima mensagem do cliente (message-in).
        Retorna o texto da última mensagem recebida ou None se timeout.
        """
        logging.info(f"Aguardando resposta (timeout={timeout}s)...")
        inicio = time.time()
        ultima_msg = None

        while (time.time() - inicio) < timeout:
            try:
                mensagens = []
                # tentativa 1: spans de mensagens de entrada
                mensagens = self.driver.find_elements(By.XPATH, '//div[contains(@class, "message-in")]//span[contains(@class,"selectable-text") or @class="selectable-text copyable-text"]')
                if not mensagens:
                    mensagens = self.driver.find_elements(By.XPATH, '//div[@data-testid="msg-in"]//div[contains(@class,"selectable-text")]//span')
                if not mensagens:
                    mensagens = self.driver.find_elements(By.XPATH, '//div[contains(@class,"message-in")]//span')

                if mensagens:
                    texto = mensagens[-1].text.strip()
                    if texto and texto != ultima_msg:
                        logging.info(f"Resposta recebida (preview): {texto[:80]}")
                        return texto
                time.sleep(poll_interval)
            except Exception as e:
                logging.debug(f"Erro verificando mensagens: {e}")
                time.sleep(1)
        logging.warning("Timeout ao aguardar resposta")
        return None

    # -------------------------
    # Montar mensagem
    # -------------------------
    def montar_mensagem_circularizacao(self, nome_cliente, notas_fiscais, nome_empresa):
        nf_lista = ""
        if isinstance(notas_fiscais, (list, tuple)):
            nf_lista = "\n".join([f"  • NF {nf.strip()}" for nf in notas_fiscais if str(nf).strip() != ""])
        else:
            # se vier string, separar por vírgula
            nf_list = [n.strip() for n in str(notas_fiscais).split(',') if n.strip() != ""]
            nf_lista = "\n".join([f"  • NF {nf}" for nf in nf_list])

        mensagem = f"""Olá, {nome_cliente}!

Sou colaborador(a) da {nome_empresa} e estamos realizando um acompanhamento das vendas do distribuidor 

Estamos entrando em contato com os clientes da COMAL COMERCIAL DE ACUMULADORES E COMPONENTES LTDA² para verificar se as vendas do Distribuidor estão ocorrendo corretamente. 

Quero reforçar que meu contato não é de cobrança!

No nosso sistema consta que vocês realizaram as seguintes compras com o distribuidor:

{nf_lista}

Vocês reconhecem essas compras e esse saldo total de baterias inservíveis (BIN)?

Por favor, responda:
SIM - Se reconhece todas as notas fiscais
NÃO - Se não reconhece alguma nota fiscal

Caso tenha alguma dúvida, estou à disposição.

Aguardo seu retorno. Obrigado(a)! 😊"""
        # Não sanitize aqui para manter originalidade; sanitizamos antes do envio.
        return mensagem

    # -------------------------
    # Processar circularização
    # -------------------------
    def processar_circularizacao(self, nome_empresa, intervalo_entre_msgs=10, timeout_resposta=300):
        df_clientes = self.carregar_clientes()
        total = len(df_clientes)
        logging.info(f"Iniciando circularização de {total} clientes...")

        for index, row in df_clientes.iterrows():
            # Inicializar variáveis ANTES do try para evitar erro de variável não definida
            nome = row.get('nome_cliente', '') if isinstance(row, dict) else row['nome_cliente']
            raw_tel = row.get('telefone', '') if isinstance(row, dict) else row['telefone']
            notas_raw = row.get('notas_fiscais', '') if isinstance(row, dict) else row.get('notas_fiscais', '')

            # normalizar
            telefone = self._sanitize_phone(raw_tel)
            # preparar lista de notas
            if isinstance(notas_raw, str):
                notas = [n.strip() for n in notas_raw.split(',') if n.strip() != ""]
            elif pd.isna(notas_raw) or notas_raw is None:
                notas = []
            else:
                # possibilidade de lista já
                try:
                    notas = list(notas_raw)
                except Exception:
                    notas = [str(notas_raw)]

            try:
                logging.info(f"\n[{index+1}/{total}] Processando: {nome} - {telefone}")

                # Validar telefone
                if not telefone:
                    logging.warning("Telefone inválido ou ausente; pulando cliente.")
                    self.respostas.append({
                        'nome_cliente': nome,
                        'telefone': raw_tel,
                        'notas_fiscais': notas_raw,
                        'status': 'ERRO - TELEFONE INVÁLIDO',
                        'resposta': None,
                        'data_hora': datetime.now()
                    })
                    continue

                # Buscar contato
                if not self.buscar_contato(telefone):
                    self.respostas.append({
                        'nome_cliente': nome,
                        'telefone': telefone,
                        'notas_fiscais': ','.join(notas),
                        'status': 'ERRO - CONTATO NÃO ENCONTRADO',
                        'resposta': None,
                        'data_hora': datetime.now()
                    })
                    # pequena espera para evitar bloqueio por ações rápidas
                    time.sleep(1)
                    continue

                # Montar e enviar mensagem
                mensagem = self.montar_mensagem_circularizacao(nome, notas, nome_empresa)
                if not self.enviar_mensagem(mensagem):
                    self.respostas.append({
                        'nome_cliente': nome,
                        'telefone': telefone,
                        'notas_fiscais': ','.join(notas),
                        'status': 'ERRO - FALHA AO ENVIAR',
                        'resposta': None,
                        'data_hora': datetime.now()
                    })
                    continue

                # Aguardar resposta
                resposta = self.aguardar_resposta(timeout=timeout_resposta)
                status = 'RESPONDIDO' if resposta else 'SEM RESPOSTA'

                self.respostas.append({
                    'nome_cliente': nome,
                    'telefone': telefone,
                    'notas_fiscais': ','.join(notas),
                    'status': status,
                    'resposta': resposta,
                    'data_hora': datetime.now()
                })

                # Intervalo entre mensagens
                if index < total - 1:
                    logging.info(f"Aguardando {intervalo_entre_msgs}s antes da próxima mensagem...")
                    time.sleep(intervalo_entre_msgs)

            except KeyboardInterrupt:
                logging.info("Interrupção manual detectada. Salvando resultados parciais...")
                self.salvar_resultados()
                raise
            except Exception as e:
                logging.error(f"Erro ao processar cliente {nome}: {e}", exc_info=True)
                try:
                    notas_str = ','.join(notas) if isinstance(notas, (list,tuple)) else str(notas)
                except Exception:
                    notas_str = str(notas_raw)
                self.respostas.append({
                    'nome_cliente': nome,
                    'telefone': telefone or raw_tel,
                    'notas_fiscais': notas_str,
                    'status': f'ERRO - {str(e)}',
                    'resposta': None,
                    'data_hora': datetime.now()
                })
                # continuar para próximo cliente

        # salvo ao final do processamento
        self.salvar_resultados()

    # -------------------------
    # Salvar resultados
    # -------------------------
    def salvar_resultados(self, arquivo_saida=None):
        if not arquivo_saida:
            arquivo_saida = f'resultados_circularizacao_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        try:
            df_resultados = pd.DataFrame(self.respostas)
            # se estiver vazio, cria DF com colunas esperadas
            if df_resultados.empty:
                df_resultados = pd.DataFrame(columns=['nome_cliente','telefone','notas_fiscais','status','resposta','data_hora'])
            df_resultados.to_excel(arquivo_saida, index=False)
            logging.info(f"Resultados salvos em: {arquivo_saida}")
        except Exception as e:
            logging.error(f"Falha ao salvar resultados: {e}", exc_info=True)

    # -------------------------
    # Fechar driver
    # -------------------------
    def fechar(self):
        if self.driver:
            try:
                logging.info("Fechando navegador...")
                self.driver.quit()
            except Exception as e:
                logging.debug(f"Erro ao fechar o driver: {e}")


# -------------------------
# Função principal
# -------------------------
def main():
    print("""
    ╔═══════════════════════════════════════════════════════════╗
    ║         BOT DE CIRCULARIZAÇÃO - AUDITORIA INTERNA         ║
    ║                      Via WhatsApp                         ║
    ╚═══════════════════════════════════════════════════════════╝
    """)

    # CONFIGURE AQUI
    PLANILHA_CLIENTES = r'C:\Users\matheus.melo\OneDrive - Acumuladores Moura SA\Documentos\Drive - Matheus Melo\Circularização\Exemplo\clientes_circularizacao.xlsx'
    NOME_EMPRESA = 'XYZ LTDA'  # altere conforme necessário
    INTERVALO_MSGS = 15         # segundos entre mensagens
    TIMEOUT_RESPOSTA = 300      # tempo máximo para aguardar resposta (segundos)
    PROFILE_PATH = os.path.join(os.getcwd(), 'whatsapp_session')
    CHROME_DRIVER_PATH = None  # opcional: ex: r'C:\drivers\chromedriver.exe'

    bot = CircularizacaoBot(
        planilha_clientes=PLANILHA_CLIENTES,
        profile_path=PROFILE_PATH,
        chrome_driver_path=CHROME_DRIVER_PATH,
        max_wait_login=120
    )

    try:
        bot.iniciar_whatsapp(headless=False)
        bot.processar_circularizacao(
            nome_empresa=NOME_EMPRESA,
            intervalo_entre_msgs=INTERVALO_MSGS,
            timeout_resposta=TIMEOUT_RESPOSTA
        )
        print("\n✓ Circularização concluída com sucesso!")
        logging.info("Processo finalizado com sucesso.")
    except KeyboardInterrupt:
        logging.info("Processo interrompido pelo usuário.")
        print("\nProcesso interrompido pelo usuário. Salvando resultados parciais...")
    except Exception as e:
        logging.error(f"Erro crítico: {e}", exc_info=True)
        print(f"\nErro crítico: {e}")
    finally:
        try:
            # garante salvamento caso algo tenha ocorrido
            bot.salvar_resultados()
        except Exception:
            pass
        bot.fechar()

if __name__ == "__main__":
    main()
