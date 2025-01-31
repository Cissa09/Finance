# 📌 Visão Geral do Projeto

Este projeto busca dados históricos de ações na **Yahoo Finance API** e gera relatórios completos no **Excel**, com análise **mensal** e **anual**.

## 🏆 Objetivo do Projeto
Criar um sistema modular para baixar e processar dados financeiros, permitindo a análise de **mínimos, máximos e variações percentuais** de um ativo ao longo do tempo.

## ⚙️ Tecnologias Utilizadas
- **C# / .NET 8.0**
- **Yahoo Finance API** (dados de mercado)
- **ClosedXML** (manipulação de arquivos Excel)
- **NodaTime** (manipulação avançada de datas)

## 📂 Estrutura do Projeto

O projeto está dividido em três classes principais:
- **`Program.cs`** → Responsável por iniciar o programa e capturar os parâmetros do usuário.
- **`ProcessaDados.cs`** → Obtém os dados do Yahoo Finance e processa a análise mensal e anual.
- **`Excel.cs`** → Manipula a criação e exportação do arquivo Excel com os dados estruturados.

---

# 🚀 Como Instalar e Rodar o Projeto

## 1️⃣ Clonar o Repositório
```sh
git clone https://github.com/Cissa09/Finance.git
cd seu-repositorio
```

## 2️⃣ Executar o Projeto
```sh
dotnet run PETR4.SA
```
🔹 **Substitua `PETR4.SA` pelo código do ativo desejado.**

## 3️⃣ Abrir o Excel Gerado
📂 O arquivo será salvo na **Área de Trabalho** do usuário com o nome:
```sh
Historico_PETR4.SA.xlsx
```
📊 O Excel terá **duas abas**:
- **Mensal** → Dados agrupados por mês
- **Anual** → Dados agrupados por ano

---

# 📖 Explicação Técnica

## 📌 `Program.cs`
Esta é a classe principal, responsável por:
✅ Capturar o ticker do ativo via argumento ou entrada no console
✅ Chamar a classe `ProcessaDados` para processar os dados
✅ Exibir mensagens informando sobre o progresso e erros

## 📌 `ProcessaDados.cs`
Responsável por:
✅ Buscar os dados históricos na **Yahoo Finance API**
✅ Processar os valores mínimos, máximos e variação
✅ Separar os dados em **análise mensal e anual**
✅ Enviar os dados formatados para a classe `Excel`

## 📌 `Excel.cs`
✅ Gera o arquivo **Excel** com os dados processados
✅ Cria **duas abas** no arquivo (`Mensal` e `Anual`)
✅ Ajusta o formato das células para melhor visualização

---

# 💡 Como Contribuir
Quer ajudar a melhorar o projeto? Siga estes passos:

## 1️⃣ Fork o Repositório
Clique em **Fork** no canto superior direito do GitHub.

## 2️⃣ Clone o Repositório
```sh
git clone https://github.com/Cissa09/Finance.git
cd seu-repositorio
```

## 3️⃣ Crie uma Nova Branch
```sh
git checkout -b minha-melhoria
```

## 4️⃣ Faça as Alterações e Envie um Pull Request
```sh
git add .
git commit -m "Melhoria no processamento dos dados"
git push origin minha-melhoria
```
🔹 Depois, vá até o **GitHub** e crie um **Pull Request**!

---

# 📜 Licença
Este projeto está sob a licença **MIT**. Sinta-se à vontade para usá-lo e melhorá-lo! 🚀

---

# 📬 Contato
Caso tenha dúvidas ou sugestões, entre em contato!
💻 [Seu GitHub](https://github.com/cissa09)  
📧 [Seu E-mail](mailto:cicero.viganon@hotmail.com)  

---

### 🎯 Agora sua Wiki está pronta para ser adicionada ao GitHub! 🚀📖
