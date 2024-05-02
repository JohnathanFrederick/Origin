~~~C

#include <Origin.h>

int num_RightNameMacht = 3;                 // O número de caracteres, a partir da direita, para comparação dos nomes das Worksheets
string name_DataLayer = "Data";             // O nome do layer, na WorksheetPage, que contém os dados a serem plotados

void PlotGraphs() {
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

void BuildGraph(vector<string> vector_MatchedNames, string name_Now){
    GraphPage gp;
    gp.Create("Origin");
    GraphLayer gl = gp.Layers();

    double min_X, max_X, min_Y, max_Y;
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
            Dataset ds_X(wks, 0);
            Dataset ds_Y(wks, 2);
            double min_X_wks = min(ds_X);
            double max_X_wks = max(ds_X);
            double min_Y_wks = min(ds_Y);
            double max_Y_wks = max(ds_Y);
            if(bool_FirstVal){
                min_X = min_X_wks;
                max_X = max_X_wks;
                min_Y = min_Y_wks;
                max_Y = max_Y_wks;
                bool_FirstVal = false;
            }else{
                if(min_X_wks < min_X) min_X = min_X_wks;
                if(max_X_wks > max_X) max_X = max_X_wks;
                if(min_Y_wks < min_Y) min_Y = min_Y_wks;
                if(max_Y_wks > max_Y) max_Y = max_Y_wks;
            }

        }else{
            out_str("A worksheet " + wksName + " não possui uma sheet com o nome " + name_DataLayer)
        }
    }
    FormatGraph(gl, min_X, max_X, min_Y, max_Y);
}

void FormatGraph(GraphLayer gl, double min_X, double max_X, double min_Y, double max_Y){
    gl.GroupPlots(0);

    // Ajuste a escala do eixo X
    Scale sc_X(gl.X);
    sc_X.From = min_X;
    sc_X.To = max_X;
    sc_X.Inc = 5;

    // Ajuste a escala do eixo Y
    Scale sc_Y(gl.Y);
    sc_Y.From = min_Y;
    sc_Y.To = max_Y;
    sc_Y.Inc = 5;

    Tree trFormat;
    // Mostre o eixo superior e seus ticks
    trFormat.Root.Axes.X.Ticks.TopTicks.show.nVal = 1;
    gl.ApplyFormat( trFormat, true, true, true );

    // Mostre o eixo direito e seus ticks
    trFormat.Root.Axes.Y.Ticks.RightTicks.show.nVal = 1;
    gl.ApplyFormat( trFormat, true, true, true );
}

~~~
