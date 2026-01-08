from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import os

# Credenciais de acesso
usuario = "matheus.melo@bateriasmoura.com"
senha = "@Mat1081042103"

# Configurações do navegador
options = webdriver.ChromeOptions()
options.add_argument("--incognito")
options.add_argument("--start-maximized")

# CONFIGURAÇÃO PARA MOSTRAR O DIÁLOGO "SALVAR COMO"
prefs = {
    "download.default_directory": "",  # String vazia para não definir diretório padrão
    "download.prompt_for_download": True,  # IMPORTANTE: Mostrar diálogo de salvamento
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    "safebrowsing.disable_download_protection": True
}
options.add_experimental_option("prefs", prefs)

# Inicia o navegador
driver = webdriver.Chrome(options=options)

# Acessa a página de login do D365
driver.get("https://bateriasmoura.operations.dynamics.com")

# Aguarda carregamento da página
time.sleep(3)

# Preenche o campo de usuário
campo_usuario = driver.find_element(By.ID, "i0116")
campo_usuario.send_keys(usuario)
campo_usuario.send_keys(Keys.RETURN)
time.sleep(3)

# Preenche o campo de senha
campo_senha = driver.find_element(By.ID, "i0118")
campo_senha.send_keys(senha)
campo_senha.send_keys(Keys.RETURN)
time.sleep(3)

# Clica em "Sim" para manter conectado (se aparecer)
try:
    botao_sim = driver.find_element(By.ID, "idSIButton9")
    botao_sim.click()
    time.sleep(3)
except:
    pass

# Função para executar o relatório de faturamento
def executar_faturamento():
    try:
        # Acessa o painel padrão
        driver.get("https://bateriasmoura.operations.dynamics.com/?cmp=r24&mi=DefaultDashboard")
        time.sleep(5)

        # Navega para o relatório de faturamento
        driver.get("https://bateriasmoura.operations.dynamics.com/?cmp=r24&mi=BillingStatementInquiryFilters")
        time.sleep(5)

        print("🔍 Iniciando preenchimento do relatório...")
        print(f"📄 URL atual: {driver.current_url}")
        print(f"📄 Título da página: {driver.title}")
        print("💡 O diálogo 'Salvar como' será aberto para você escolher onde salvar o arquivo")

        # *PREENCHER CAMPO DA EMPRESA*
        try:
            campo_empresa = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'Company') or contains(@name, 'Company')]"))
            )
            campo_empresa.clear()
            campo_empresa.send_keys("R241")
            print("✅ Campo empresa preenchido com 'R241'")
            time.sleep(3)  # Aguarda lista carregar
        except Exception as e:
            print(f"❌ Erro ao preencher campo empresa: {e}")
            return

        # *CLICAR NO CHECKBOX/FLAG*
        try:
            print("🔍 Procurando checkbox/flag...")
            
            # Estratégias para encontrar o checkbox
            seletores_checkbox = [
                (By.XPATH, "//div[contains(@class, 'dyn-hoverMarkingColumn')]"),
                (By.XPATH, "//div[@class='dyn-hoverMarkingColumn']"),
                (By.XPATH, "//div[contains(@class, 'public_fixedDataTableCell_cellContent')]//div[contains(@class, 'dyn-hoverMarkingColumn')]"),
                (By.XPATH, "//div[contains(@class, 'hoverMarkingColumn')]"),
                (By.XPATH, "//div[contains(@class, 'markingColumn')]"),
            ]
            
            checkbox_encontrado = False
            for seletor_type, seletor_value in seletores_checkbox:
                if not checkbox_encontrado:
                    try:
                        print(f"  Tentando encontrar checkbox com: {seletor_value}")
                        checkbox = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((seletor_type, seletor_value))
                        )
                        
                        # Verifica se o elemento está visível e clicável
                        if checkbox.is_displayed() and checkbox.is_enabled():
                            checkbox.click()
                            print(f"✅ Checkbox clicado com sucesso (usando {seletor_value})")
                            checkbox_encontrado = True
                            time.sleep(2)
                            break
                    except Exception as e:
                        print(f"  ❌ Não encontrado com {seletor_value}: {e}")
                        continue
            
            if not checkbox_encontrado:
                print("⚠️  Checkbox não encontrado com os seletores padrão. Tentando estratégia alternativa...")
                
                # Estratégia alternativa: procurar por elementos que possam ser checkboxes
                elementos_suspeitos = driver.find_elements(By.XPATH, "//div[contains(@class, 'checkbox') or contains(@class, 'check') or contains(@class, 'select') or contains(@class, 'mark')]")
                print(f"🔍 Encontrados {len(elementos_suspeitos)} elementos suspeitos de serem checkboxes")
                
                for i, elemento in enumerate(elementos_suspeitos):
                    try:
                        outer_html = elemento.get_attribute('outerHTML')
                        class_attr = elemento.get_attribute('class')
                        print(f"  Elemento {i+1}: class='{class_attr}'")
                        
                        # Se parece com um elemento de marcação/seleção, tenta clicar
                        if 'dyn-hoverMarkingColumn' in class_attr or 'hoverMarking' in class_attr or 'marking' in class_attr:
                            elemento.click()
                            print(f"✅ Checkbox clicado (elemento {i+1} com class '{class_attr}')")
                            checkbox_encontrado = True
                            time.sleep(2)
                            break
                    except Exception as e:
                        print(f"  ❌ Não foi possível clicar no elemento {i+1}: {e}")
                        continue
            
            if not checkbox_encontrado:
                print("❌ Não foi possível encontrar e clicar no checkbox. Continuando sem marcar o checkbox...")
                
        except Exception as e:
            print(f"❌ Erro ao tentar clicar no checkbox: {e}")

        # *CLICAR EM SELECIONAR*
        try:
            botao_selecionar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Selecionar')]"))
            )
            botao_selecionar.click()
            print("✅ Botão 'Selecionar' clicado (por qualquer elemento com texto)")
            time.sleep(3)
        except Exception as e:
            print(f"❌ Todas as tentativas falharam: {e}")
            return

        # *PREENCHER DATAS - COM MÚLTIPLAS TENTATIVAS*
        print("🔍 Procurando campos de data...")
        
        # Estratégias para encontrar a data inicial
        seletores_data_inicio = [
            (By.ID, "StartDate"),
            (By.XPATH, "//input[contains(@id, 'dateFrom') or contains(@name, 'dateFrom')]"),
            (By.XPATH, "//input[contains(@id, 'StartDate') or contains(@name, 'StartDate')]"),
            (By.XPATH, "//input[contains(@placeholder, 'data') or contains(@aria-label, 'data')]"),
            (By.XPATH, "//input[@type='date' or @type='text']")
        ]

        data_preenchida = False
        for seletor_type, seletor_value in seletores_data_inicio:
            if not data_preenchida:
                try:
                    data_inicio = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((seletor_type, seletor_value))
                    )
                    data_inicio.clear()
                    data_inicio.send_keys("01/09/2025")
                    print(f"✅ Data início preenchida: 01/09/2025 (usando {seletor_value})")
                    data_preenchida = True
                    break
                except:
                    continue

        if not data_preenchida:
            print("❌ Não foi possível encontrar o campo de data início")

        # Estratégias para encontrar a data final
        seletores_data_fim = [
            (By.ID, "EndDate"),
            (By.XPATH, "//input[contains(@id, 'dateTo') or contains(@name, 'dateTo')]"),
            (By.XPATH, "//input[contains(@id, 'EndDate') or contains(@name, 'EndDate')]"),
            (By.XPATH, "(//input[@type='date' or @type='text'])[2]")
        ]

        data_fim_preenchida = False
        for seletor_type, seletor_value in seletores_data_fim:
            if not data_fim_preenchida:
                try:
                    data_fim = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((seletor_type, seletor_value))
                    )
                    data_fim.clear()
                    data_fim.send_keys("30/09/2025")
                    print(f"✅ Data fim preenchida: 30/09/2025 (usando {seletor_value})")
                    data_fim_preenchida = True
                    break
                except:
                    continue

        if not data_fim_preenchida:
            print("❌ Não foi possível encontrar o campo de data fim")

        time.sleep(2)

        # *CLICAR EM OK - COM MÚLTIPLAS TENTATIVAS MELHORADAS*
        print("🔍 Procurando botão OK...")
        
        seletores_ok = [
            (By.ID, "billingstatementinquiryfilters_4_FormCommandButtonOK_label"),
            (By.XPATH, "//*[@id='billingstatementinquiryfilters_4_FormCommandButtonOK_label']"),
            (By.XPATH, "//button[contains(text(), 'OK')]"),
            (By.XPATH, "//input[@value='OK']"),
            (By.XPATH, "//*[text()='OK']"),
            (By.XPATH, "//button[contains(@class, 'ok') or contains(@class, 'OK')]"),
            (By.XPATH, "//input[contains(@class, 'ok') or contains(@class, 'OK')]"),
            (By.XPATH, "//a[contains(text(), 'OK')]"),
            (By.XPATH, "//div[contains(text(), 'OK')]"),
            (By.XPATH, "//span[contains(text(), 'OK')]"),
            (By.XPATH, "//*[contains(@title, 'OK') or contains(@aria-label, 'OK')]"),
            # Tentativas com números diferentes (o número pode mudar)
            (By.ID, "billingstatementinquiryfilters_3_FormCommandButtonOK_label"),
            (By.ID, "billingstatementinquiryfilters_5_FormCommandButtonOK_label"),
            (By.ID, "billingstatementinquiryfilters_6_FormCommandButtonOK_label"),
        ]

        ok_clicado = False
        for seletor_type, seletor_value in seletores_ok:
            if not ok_clicado:
                try:
                    print(f"  Tentando: {seletor_value}")
                    botao_ok = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((seletor_type, seletor_value))
                    )
                    botao_ok.click()
                    print(f"✅ Botão OK clicado (usando {seletor_value})")
                    ok_clicado = True
                    
                    # **AGUARDAR O RELATÓRIO CARREGAR COMPLETAMENTE - ESPERA FIXA DE 3 MINUTOS**
                    print("⏳ Aguardando o relatório carregar (aguardando 3 minutos fixos)...")
                    
                    # Timer para mostrar progresso
                    start_time = time.time()
                    total_wait_time = 180  # 3 minutos em segundos
                    
                    # Aguarda exatamente 3 minutos (180 segundos) - tempo fixo
                    for seconds_passed in range(total_wait_time):
                        if seconds_passed % 30 == 0:  # Mostra progresso a cada 30 segundos
                            minutes_passed = seconds_passed // 60
                            seconds_remaining = total_wait_time - seconds_passed
                            print(f"⏰ Aguardando... {minutes_passed}min {seconds_passed % 60}s passados | {seconds_remaining // 60}min {seconds_remaining % 60}s restantes")
                        
                        time.sleep(1)
                    
                    elapsed_time = time.time() - start_time
                    print(f"✅ Tempo de espera concluído: {elapsed_time:.1f} segundos")
                    
                    # Aguarda um tempo adicional para garantir carregamento completo
                    time.sleep(10)
                    print("✅ Aguardando tempo adicional de segurança")
                    
                    break
                except Exception as e:
                    continue

        if not ok_clicado:
            print("❌ Não foi possível encontrar e clicar no botão OK")
            print("💡 Tente verificar manualmente qual é o ID/texto do botão OK na página")
            return

        # *EXPORTAR RELATÓRIO - CLIQUE DIREITO NO ELEMENTO ESPECÍFICO*
        try:
            print("🔍 Preparando para exportar o relatório...")
            
            # Aguardar um pouco mais para garantir que tudo está carregado
            time.sleep(5)
            
            # USAR O XPATH ESPECÍFICO FORNECIDO
            xpath_especifico = "//*[@id='GridCell-0-BillingStatementInquiry_Name']/div/div/div"
            
            print(f"🔍 Procurando elemento específico com XPath: {xpath_especifico}")
            
            # Aguardar e encontrar o elemento específico
            elemento_alvo = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.XPATH, xpath_especifico))
            )
            
            print("✅ Elemento específico encontrado e está clicável")
            
            # Verificar se o elemento está visível
            if elemento_alvo.is_displayed():
                print("✅ Elemento está visível na tela")
            else:
                print("⚠️  Elemento não está visível, mas tentando mesmo assim...")
            
            # Criar ActionChains para clique direito
            actions = ActionChains(driver)
            
            # Clicar com botão direito no elemento específico
            print("🖱️  Clicando com botão direito no elemento específico...")
            actions.context_click(elemento_alvo).perform()
            time.sleep(3)
            
            # Procurar e clicar na opção "Exportar todas as linhas"
            print("🔍 Procurando opção 'Exportar todas as linhas'...")
            
            seletores_exportar = [
                (By.XPATH, "//*[contains(text(), 'Exportar todas as linhas')]"),
                (By.XPATH, "//*[contains(text(), 'Export all rows')]"),
                (By.XPATH, "//*[contains(text(), 'Exportar') and contains(text(), 'linhas')]"),
                (By.XPATH, "//*[contains(text(), 'Export') and contains(text(), 'rows')]"),
                (By.XPATH, "//div[contains(@class, 'context')]//*[contains(text(), 'Exportar')]"),
                (By.XPATH, "//div[contains(@class, 'menu')]//*[contains(text(), 'Exportar')]"),
                (By.XPATH, "//*[@role='menu']//*[contains(text(), 'Exportar')]"),
                (By.XPATH, "//*[@role='menuitem']//*[contains(text(), 'Exportar')]"),
            ]
            
            exportar_clicado = False
            for seletor_type, seletor_value in seletores_exportar:
                if not exportar_clicado:
                    try:
                        print(f"  Tentando opção: {seletor_value}")
                        opcao_exportar = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((seletor_type, seletor_value))
                        )
                        opcao_exportar.click()
                        print(f"✅ Opção 'Exportar todas as linhas' clicada (usando {seletor_value})")
                        exportar_clicado = True
                        
                        # Aguardar a janela de download aparecer
                        print("⏳ Aguardando janela de download aparecer...")
                        time.sleep(10)
                        
                        # *CLICAR NO BOTÃO BAIXAR - COM O NOVO XPATH*
                        print("🔍 Procurando botão 'Baixar'...")
                        xpath_botao_baixar = "//*[@id='DocuFileSaveDialog_5_DownloadButton']"
                        
                        try:
                            botao_baixar = WebDriverWait(driver, 15).until(
                                EC.element_to_be_clickable((By.XPATH, xpath_botao_baixar))
                            )
                            botao_baixar.click()
                            print("✅ Botão 'Baixar' clicado com sucesso!")
                            
                            print("🎯" + "="*70)
                            print("🎯 DIÁLOGO 'SALVAR COMO' ABERTO!")
                            print("🎯" + "="*70)
                            print("💡 Agora você pode escolher manualmente onde salvar o arquivo.")
                            print("📁 Selecione a pasta desejada e clique em 'Salvar'.")
                            print("⏰ O navegador NÃO será fechado automaticamente.")
                            print("")
                            print("🔄 Quando terminar de salvar o arquivo:")
                            print("1. Feche a janela do diálogo 'Salvar como'")
                            print("2. Volte para este terminal")
                            print("3. Pressione ENTER para fechar o navegador")
                            print("")
                            
                            # **ESPERA MANUAL - O USUÁRIO DECIDE QUANDO FECHAR**
                            input("👉 Pressione ENTER para fechar o navegador...")
                            
                            print("✅ Salvamento manual concluído pelo usuário.")
                            
                        except Exception as e:
                            print(f"❌ Erro ao clicar no botão 'Baixar': {e}")
                            print("💡 Tentando estratégias alternativas para o botão Baixar...")
                            
                            # Estratégias alternativas para o botão Baixar
                            seletores_baixar_alternativos = [
                                (By.XPATH, "//*[@id='DocuFileSaveDialog_5_DownloadButton_label']"),
                                (By.XPATH, "//button[contains(text(), 'Baixar')]"),
                                (By.XPATH, "//input[@value='Baixar']"),
                                (By.XPATH, "//*[contains(text(), 'Baixar')]"),
                                (By.XPATH, "//button[contains(text(), 'Download')]"),
                                (By.XPATH, "//input[@value='Download']"),
                                (By.XPATH, "//*[contains(text(), 'Download')]"),
                            ]
                            
                            baixar_clicado = False
                            for seletor_type_baixar, seletor_value_baixar in seletores_baixar_alternativos:
                                if not baixar_clicado:
                                    try:
                                        print(f"  Tentando botão alternativo: {seletor_value_baixar}")
                                        botao_alt = WebDriverWait(driver, 5).until(
                                            EC.element_to_be_clickable((seletor_type_baixar, seletor_value_baixar))
                                        )
                                        botao_alt.click()
                                        print(f"✅ Botão de download clicado (alternativo: {seletor_value_baixar})")
                                        baixar_clicado = True
                                        
                                        print("🎯" + "="*70)
                                        print("🎯 DIÁLOGO 'SALVAR COMO' ABERTO!")
                                        print("🎯" + "="*70)
                                        print("💡 Agora você pode escolher manualmente onde salvar o arquivo.")
                                        print("📁 Selecione a pasta desejada e clique em 'Salvar'.")
                                        print("⏰ O navegador NÃO será fechado automaticamente.")
                                        print("")
                                        print("🔄 Quando terminar de salvar o arquivo:")
                                        print("1. Feche a janela do diálogo 'Salvar como'")
                                        print("2. Volte para este terminal")
                                        print("3. Pressione ENTER para fechar o navegador")
                                        print("")
                                        
                                        # **ESPERA MANUAL - O USUÁRIO DECIDE QUANDO FECHAR**
                                        input("👉 Pressione ENTER para fechar o navegador...")
                                        
                                        print("✅ Salvamento manual concluído pelo usuário.")
                                        break
                                    except Exception as alt_e:
                                        continue
                            
                            if not baixar_clicado:
                                print("❌ Não foi possível encontrar o botão de download")
                        
                        break
                    except Exception as e:
                        print(f"  ❌ Opção não encontrada com {seletor_value}: {e}")
                        continue
            
            if not exportar_clicado:
                print("❌ Não foi possível encontrar a opção 'Exportar todas as linhas'")
                print("💡 Tentando estratégia alternativa...")
                
                # Estratégia alternativa: procurar por qualquer opção de exportação
                try:
                    opcoes_exportacao = driver.find_elements(By.XPATH, "//*[contains(text(), 'Exportar')] | //*[contains(text(), 'Export')]")
                    print(f"🔍 Encontradas {len(opcoes_exportacao)} opções de exportação")
                    
                    for i, opcao in enumerate(opcoes_exportacao):
                        try:
                            if opcao.is_displayed():
                                texto_opcao = opcao.text
                                print(f"  Opção {i+1}: '{texto_opcao}'")
                                if 'exportar' in texto_opcao.lower() or 'export' in texto_opcao.lower():
                                    opcao.click()
                                    print(f"✅ Opção de exportação clicada: '{texto_opcao}'")
                                    
                                    # Aguardar a janela de download aparecer
                                    time.sleep(10)
                                    
                                    # Tentar clicar no botão Baixar após exportação alternativa
                                    try:
                                        botao_baixar = WebDriverWait(driver, 15).until(
                                            EC.element_to_be_clickable((By.XPATH, "//*[@id='DocuFileSaveDialog_5_DownloadButton']"))
                                        )
                                        botao_baixar.click()
                                        print("✅ Botão 'Baixar' clicado após exportação alternativa")
                                        
                                        print("🎯" + "="*70)
                                        print("🎯 DIÁLOGO 'SALVAR COMO' ABERTO!")
                                        print("🎯" + "="*70)
                                        print("💡 Agora você pode escolher manualmente onde salvar o arquivo.")
                                        print("📁 Selecione a pasta desejada e clique em 'Salvar'.")
                                        print("⏰ O navegador NÃO será fechado automaticamente.")
                                        print("")
                                        print("🔄 Quando terminar de salvar o arquivo:")
                                        print("1. Feche a janela do diálogo 'Salvar como'")
                                        print("2. Volte para este terminal")
                                        print("3. Pressione ENTER para fechar o navegador")
                                        print("")
                                        
                                        # **ESPERA MANUAL - O USUÁRIO DECIDE QUANDO FECHAR**
                                        input("👉 Pressione ENTER para fechar o navegador...")
                                        
                                        print("✅ Salvamento manual concluído pelo usuário.")
                                        
                                    except Exception as e:
                                        print(f"❌ Não foi possível clicar no botão Baixar: {e}")
                                    
                                    exportar_clicado = True
                                    break
                        except Exception as e:
                            print(f"  ❌ Não foi possível clicar na opção {i+1}: {e}")
                            continue
                except Exception as e:
                    print(f"❌ Estratégia alternativa também falhou: {e}")

        except Exception as e:
            print(f"❌ Erro na exportação: {e}")

    except Exception as e:
        print("❌ Erro geral na execução:", str(e))

# Executa a função
executar_faturamento()

# Encerra o navegador
print("")
print("🔚 Fechando navegador...")
driver.quit()
print("✅ Navegador fechado. Processo concluído!")
