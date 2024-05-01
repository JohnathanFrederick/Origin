~~~C

#include <Origin.h>

// Plota um conjunto de gráficos a partir de um conjunto de dados. Os dados são agrupados em cada gráfico com base nos nomes de cada Worksheet.
void Teste_Plot() {
    int num_RightNameMacht = 3;                 // O número de caracteres, a partir da direita, para comparação dos nomes das Worksheets
    string name_DataLayer = "Data";             // O nome do layer, na WorksheetPage, que contém os dados a serem plotados

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

            // Crie o GraphPage e o GraphLayer
            vector_AlreadyCreated.Add(name_Now);
            GraphPage gp;
            gp.Create("Origin");
            GraphLayer gl = gp.Layers();

            // Processe as Worksheets com nomes correpondentes
            double min_X, max_X;
            bool bool_FirstVal = true;
            for(int j = 0; j < vector_MatchedNames.GetSize(); j++){
                string wksName = vector_MatchedNames[j];

                // Obtenha a Worksheet com o nome correspondente
                out_str("Adicionando os dados de: " + wksName + " ao grupo de: " + name_Now);
                WorksheetPage wksPage = Project.WorksheetPages(wksName);

                // Obtenha os dados
                DataRange dr;
                Worksheet wks = wksPage.Layers(name_DataLayer);
                if(wks){
                    // Construa o DataRange e o adicione ao Layer;
                    dr.Add(wks, 0, "X");
                    dr.Add(wks, 2, "Y");
                    gl.AddPlot(dr, IDM_PLOT_LINE);

                    // Atualize os máximos e mínimos
                    Dataset ds_X(wks,0);
                    double min_X_wks = min(ds_X);
                    double max_X_wks = max(ds_X);
                    if(bool_FirstVal){
                        min_X = min_X_wks;
                        max_X = max_X_wks;
                        bool_FirstVal = false;
                    }else{
                        if(min_X_wks < min_X) min_X = min_X_wks;
                        if(max_X_wks > max_X) max_X = max_X_wks;
                    }

                }else{
                    out_str("A worksheet " + wksName + " não possui uma sheet com o nome " + name_DataLayer)
                }
            }
            
            // Personalize o Layer
            gl.GroupPlots(0);
            Scale sc(gl.X);
            sc.From = min_X;
            sc.To = max_X;
            sc.Inc = 1;
        }
    }
}


~~~
