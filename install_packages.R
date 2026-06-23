# Script para instalar todas as dependências necessárias
# Este arquivo será usado pelo Dockerfile

# Lista de pacotes necessários
packages <- c(
  "shiny",
  "DT", 
  "tm",
  "openxlsx",
  "shinycustomloader",
  "markdown",
  "rsconnect",
  "httr"
)

# Instalar pacotes
install.packages(packages, repos = "https://cran.rstudio.com/")

# Verificar se todos foram instalados
for(pkg in packages) {
  if(!require(pkg, character.only = TRUE)) {
    stop(paste("Falha ao instalar o pacote:", pkg))
  }
}

cat("Todas as dependências foram instaladas com sucesso!\n")