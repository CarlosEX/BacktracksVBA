![analise_combinatoria_vba](https://github.com/user-attachments/assets/5a2dd7f6-ef13-498c-b4cd-046da1056339)

# 📦 Análise Combinatória no VBA para Otimização Logística

## 🔍 Sobre o Projeto
Este projeto utiliza algoritmos de análise combinatória e técnicas de programação em **VBA** para resolver problemas complexos de otimização logística. O objetivo principal é organizar produtos em bandejas de forma eficiente, considerando restrições como a quantidade de gavetas disponíveis e combinações inteligentes.

---

## 🎯 Funcionalidades
1. **Distribuição Inteligente de Produtos**:
   - Identifica quando um produto possui mais de uma gaveta e faz o bloqueio adequado.
   - Gerencia e identifica gavetas disponíveis para maximizar o uso eficiente do espaço.

2. **Combinações Otimizadas**:
   - Encontra combinações inteligentes de produtos/gavetas como 9+1, 8+2, 7+3, etc.
   - Maximiza o uso de gavetas disponíveis, minimizando espaços vazios.

3. **Algoritmos de Backtracking**:
   - Resolve problemas de combinação utilizando técnicas como:
     - `BacktrackExactMatchAnalyzer`: Identifica mercadorias cuja quantidade de gavetas ajustadas é exatamente igual à meta predefinida.
     - `BacktrackAnalyzerMetaPair`: Gera pares otimizados para atender ao número de gavetas desejado.
     - `BacktrackAnalyzerCombination`: Procura combinações que somem ao total necessário.

4. **Integração com Solver do Excel**:
   - Explora as capacidades de modelagem de problemas do Solver, mas supera suas limitações ao implementar algoritmos personalizados para soluções mais rápidas e inteligentes.

---

## 🛠 Tecnologias Utilizadas
- **VBA (Visual Basic for Applications)**: Para lógica e automação no Excel.
- **Microsoft Excel**: Como interface para entrada de dados e exibição de resultados.
- **Solver do Excel**: Para modelagem inicial e resolução de problemas de programação linear.

---

## 📊 Exemplos de Aplicação
### 1. **Algoritmo BacktrackExactMatchAnalyzer**
Identifica produtos que atendem exatamente a uma meta predefinida de gavetas (exemplo: M = 10).  
**Resultado**:  
- Mercadorias marcadas com um "X" e removidas da lista.  

### 2. **Algoritmo BacktrackAnalyzerMetaPair**
Composição de pares que juntos completam a quantidade de gavetas desejada.  
**Resultado**:  
- Otimização dos pares, começando do menor valor para cima.

### 3. **Algoritmo BacktrackAnalyzerCombination**
Criação de combinações que somam ao total necessário para a bandeja.  
**Exemplo**:  
- Item com 7 gavetas procura itens com 3 gavetas para atender à meta de 10.

---

## 🚀 Como Usar
1. **Configuração Inicial**:
   - Abra o arquivo Excel contendo o código VBA.
   - Certifique-se de habilitar macros no Excel.

2. **Execução**:
   - Insira os dados de produtos e suas respectivas gavetas.
   - Escolha o algoritmo desejado para executar a otimização.

3. **Resultados**:
   - Os resultados serão exibidos diretamente na planilha, com combinações marcadas ou produtos agrupados conforme os critérios.

---

## 🤝 Contribuição
Sinta-se à vontade para contribuir com melhorias no projeto, relatando problemas ou sugerindo novos algoritmos. Para colaborar:
1. Faça um fork deste repositório.
2. Crie um branch para sua feature (`git checkout -b minha-feature`).
3. Envie um pull request.

---

## 📝 Licença
Este projeto está sob a licença MIT. Consulte o arquivo `LICENSE` para mais informações.

---

## 📬 Contato
Desenvolvido por **Carlos Antonio**  
💼 **LinkedIn**: [Carlos Antonio](https://www.linkedin.com/in/carlosantonio/)  
📧 **Email**: [carlosantonio@email.com](mailto:carlosantonio@email.com)

*“A tecnologia abre portas para o futuro, trazendo possibilidades infinitas e transformando sonhos em realidade.” 🚀*
