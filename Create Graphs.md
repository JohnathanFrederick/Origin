~~~C

#include <Origin.h>

void Teste() {
    vector<string> vector_Names;
    vector<string> vector_AlreadyNames;

    // Obtenha o nome de todas as WorkSheets no Projeto
    foreach(WorksheetPage wksPage in Project.WorksheetPages) {
        string name_Page = wksPage.GetName().Right(3);
        vector_Names.Add(name_Page);
    }

    // Percorra cada nome
    for(int i = 0; i < vector_Names.GetSize(); i++) {
        string name_Now = vector_Names[i];

        // Verifique se esse nome já foi gerado os gráficos
        bool alreadyPlotted = false;
        for(int j = 0; j < vector_AlreadyNames.GetSize(); j++) {
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
                if(wksPage.GetName().Right(3) == name_Now) {
                	out_str("Adicionando os dados de: " + wksPage.GetName() + " ao grupo de: " + name_Now);
                    DataRange dr;
                    Worksheet wks = wksPage.Layers();
                    dr.Add(wks, 0, "X");
                    dr.Add(wks, 2, "Y");
                    gl.AddPlot(dr, IDM_PLOT_LINE);
                }
            }
        }
    }
}

~~~
