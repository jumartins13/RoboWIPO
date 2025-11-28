import sys, os, time, subprocess

print("🚀 Iniciando robô WIPO...")
print()

# Instalar dependências
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
except ImportError:
    print("📦 Instalando Selenium...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "selenium==4.15.2"])
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options

print("🌐 Abrindo Chrome...")
print()

# Configurar Chrome
options = Options()
options.add_argument('--start-maximized')
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option('excludeSwitches', ['enable-logging'])
prefs = {
    "download.default_directory": os.path.expanduser("~/Downloads"),
    "download.prompt_for_download": False,
}
options.add_experimental_option("prefs", prefs)

driver = None

try:
    # Iniciar Chrome
    driver = webdriver.Chrome(options=options)

    print("✅ Chrome aberto!")
    print()

    # Acessar WIPO
    print("📡 Acessando site WIPO...")
    driver.get("https://www3.wipo.int/madrid/monitor/en/")
    time.sleep(6)
    print("✅ Site carregado!")
    print()

    # Modo avançado
    print("🔧 Ativando modo avançado...")
    try:
        adv = driver.find_element(By.ID, "advancedModeLink")
        adv.click()
        time.sleep(1.5)
        print("✅ Modo avançado ativado!")
        print()
    except Exception as e:
        print(f"⚠️  {e}")

    # Clicar no campo ENN
    print("📝 Clicando no campo ENN...")
    try:
        inp = driver.find_element(By.ID, "TRANSACT_input")
        inp.click()
        time.sleep(0.5)
        print("✅ Campo ENN selecionado!")
        print()
    except Exception as e:
        print(f"⚠️  {e}")

    # Pedir para digitar
    print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print("⏳ DIGITE O ENN NO NAVEGADOR AGORA!")
    print()
    print("   O campo já está selecionado e pronto.")
    print("   Você tem 15 SEGUNDOS para digitar.")
    print()
    print("   Após digitar, aguarde...")
    print("   O robô continuará automaticamente!")
    print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print()

    # Contagem regressiva
    for i in range(15, 0, -1):
        print(f"   ⏱️  {i} segundos restantes...", end='\r')
        time.sleep(1)

    print("\n")
    print("✅ Tempo esgotado! Continuando automaticamente...")
    print()

    # Clicar em Search
    print("🔍 Clicando em Search...")
    try:
        found = False
        spans = driver.find_elements(By.CSS_SELECTOR, "span.ui-button-text")
        for span in spans:
            if "search" in span.text.lower():
                span.click()
                found = True
                break

        if found:
            print("✅ Search clicado!")
        else:
            print("⚠️  Botão search não encontrado")
        print()
    except Exception as e:
        print(f"⚠️  {e}")

    # Aguardar resultados
    print("⏳ Aguardando resultados aparecerem...")
    time.sleep(6)
    print("✅ Resultados carregados!")
    print()
    print("👀 Você pode ver os resultados na tela!")
    print()

    # Pequena pausa para ver os resultados
    print("⏳ Aguardando 3 segundos para você ver os resultados...")
    time.sleep(3)

    # Gerar relatório
    print("📊 Gerando relatório...")
    try:
        gen = driver.find_element(By.ID, "generate_report")
        gen.click()
        time.sleep(2)
        print("✅ Relatório gerado!")
        print()
    except Exception as e:
        print(f"⚠️  {e}")

    # Download XLS
    print("📥 Baixando arquivo Excel...")
    try:
        dl = driver.find_element(By.ID, "download_link_xls")
        dl.click()
        print("✅ Download iniciado!")
        print()
    except Exception as e:
        print(f"⚠️  {e}")

    time.sleep(3)

    print()
    print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print("  ✅ PROCESSO CONCLUÍDO COM SUCESSO!")
    print("  📥 Arquivo Excel salvo em: ~/Downloads")
    print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print()

except Exception as e:
    print()
    print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print("  ❌ ERRO")
    print(f"  {e}")
    print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print()

finally:
    if driver:
        print("🔒 Fechando navegador em 5 segundos...")
        time.sleep(5)
        driver.quit()
        print("✅ Navegador fechado!")
        print()

input("Pressione Enter para sair...")
