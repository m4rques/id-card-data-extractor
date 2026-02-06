# ID Card Data Extractor

Uma solução de automação em Python desenvolvida para otimizar o fluxo de confecção de crachás institucionais. O script integra-se ao Microsoft Outlook para processar solicitações, extrair dados cadastrais e organizar anexos de forma automática.

## 🚀 Funcionalidades

- **Integração com Outlook:** Varredura automática de e-mails não lidos em contas específicas.
- **Extração Inteligente (Regex):** Identificação de Nome, Matrícula e Secretaria diretamente do corpo do e-mail.
- **Tratamento de Dados:** Normalização de texto, remoção de caracteres especiais e limpeza de nomes para compatibilidade com o sistema de arquivos.
- **Gestão de Anexos:** Identifica, baixa e renomeia automaticamente fotos (JPG, PNG, etc.) utilizando o número da matrícula para evitar erros de identificação.
- **Registro em Log:** Geração de um arquivo CSV consolidado com todos os dados processados para fácil importação em softwares de design de crachás.

## 🛠️ Tecnologias Utilizadas

- **Python 3.9.13**
- **pywin32 (MAPI):** Para comunicação nativa com a API do Microsoft Outlook.
- **Regular Expressions (re):** Para parsing de texto estruturado e não estruturado.
- **Pathlib:** Para manipulação de diretórios e segurança de caminhos de arquivo.
- **CSV:** Para persistência de dados.

## 📋 Pré-requisitos

Para rodar este projeto, você precisará ter o Microsoft Outlook instalado e configurado na máquina, além das seguintes dependências:

```bash
pip install pywin32
