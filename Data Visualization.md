# Produção de Gráficos em Massa

Esta seção mostra algumas funções que plotam gráficos em massa no Origin, com Origin C.
A função abaixo agrupa as WorksheetPages no arquivo Origin aberto pelo ShortName, comparando um número inteiro, parametrizado, de caracteres a partir do final do ShortName. Após este agrupamento, a função aplica uma outra função para construção de gráficos para cada conjunto de nome obtido.

~~~C
void PlotGraphs() {
  // O número de caracteres, a partir da direita, para comparação dos nomes das Worksheets
  int num_RightNameMacht = 3;
  
  // Obtenha o nome de todas as WorkSheets no Projeto para comparação...
  out_str("Obtendo o nome de todas as worksheets no Projeto para comparação...");
  vector<string> vector_WksNames;
  foreach(WorksheetPage wksPage in Project.WorksheetPages){
    string name_Page = wksPage.GetName();
    string name_ToMatch = name_Page.Right(num_RightNameMacht);
    vector_WksNames.Add(name_ToMatch);
    out_str(name_Page);
  }

  // Crie os gráficos com base no nome
  vector<string> vector_AlreadyCreated;
  for(int i = 0; i < vector_WksNames.GetSize(); i++){
    string name_Now = vector_WksNames[i];

    // Verifique se o gráfico correspondente a este nome já foi gerado
    bool alreadyPlotted = false;
    for(int j = 0; j < vector_AlreadyCreated.GetSize(); j++){
      if(vector_AlreadyCreated[j] == name_Now) {
        alreadyPlotted = true;
        break;
      }
    }

    // Se ainda não foi gerado os gráficos, gere os gráficos
    if(!alreadyPlotted) {
      // Obtenha, em ordem alfabética, os nomes de cada Worksheet que corresponde ao grupo atual
      vector<string> vector_MatchedNames;
      foreach(WorksheetPage wksPage in Project.WorksheetPages) {
        if(wksPage.GetName().Right(num_RightNameMacht) == name_Now) {
          vector_MatchedNames.Add(wksPage.GetName());
        }
      }
      vector_MatchedNames.Sort();
      vector_AlreadyCreated.Add(name_Now);

      // Construa o gráfico
      BuildGraph(vector_MatchedNames, name_Now);
    }
  }
}
~~~
