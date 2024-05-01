~~~C


#include <Origin.h>

void Teste_Plot() {
    int num_Right = 3;      // Número de caracteres, a partir da direita, para comparação dos nomes obtidos de cada Worksheet

    // Obtenha o nome de todas as WorkSheets no Projeto
    vector<string> vector_Names;
    out_str("Obtendo o nome de todas as worksheets no Projeto...");
    foreach(WorksheetPage wksPage in Project.WorksheetPages){
        string name_Page = wksPage.GetName().Right(num_Right);
        vector_Names.Add(name_Page);
        out_str(name_Page);
    }

    // Crie os gráficos com base no nome
    vector<string> vector_AlreadyNames;
    for(int i = 0; i < vector_Names.GetSize(); i++){
        string name_Now = vector_Names[i];

        // Verifique se esse nome já foi gerado os gráficos
        bool alreadyPlotted = false;
        for(int j = 0; j < vector_AlreadyNames.GetSize(); j++){
            if(vector_AlreadyNames[j] == name_Now) {
                alreadyPlotted = true;
                break;
            }
        }

        // Se ainda não foi gerado os gráficos, gere os gráficos
        if(!alreadyPlotted) {
            vector_AlreadyNames.Add(name_Now);
            GraphPage gp;
            gp.Create("Origin");
            GraphLayer gl = gp.Layers();

            // Construa o DataRange para cada WorksheetPage que corresponde ao grupo atual
            foreach(WorksheetPage wksPage in Project.WorksheetPages) {
                if(wksPage.GetName().Right(num_Right) == name_Now) {
                	out_str("Adicionando os dados de: " + wksPage.GetName() + " ao grupo de: " + name_Now);
                    DataRange dr;
                    Worksheet wks = wksPage.Layers("Data");
                    if(wks){
                        dr.Add(wks, 0, "X");
                        dr.Add(wks, 2, "Y");
                        gl.AddPlot(dr, IDM_PLOT_LINE);
                    }else{
                        out_str("A worksheet " + wksPage.GetName() + " não possui uma sheet com o nome 'Data'.")
                    }
                }
            }
            gl.GroupPlots(0);
        }
    }
}


~~~
