# Produção de Gráficos em Massa
Esta seção trata da produção de gráficos em massa no Origin C. Para tanto todas as funções abaixo são necessárias.

## Agrupamento das WorksheetsPages pelo ShortName
A função abaixo agrupa as WorksheetPages no arquivo Origin aberto pelo ShortName, comparando um número inteiro, parametrizado, de caracteres a partir do final do ShortName. Após este agrupamento, a função aplica uma outra função para construção de gráficos para cada conjunto de nome obtido.

~~~C
/**
 * A Função obtém uma quantidade parametrizada, a partir da direita, de strings únicas a partir do nome de todas as WorksheetPages presente no Projeto. A partir disso a função constrói os gráficos correspondentes a cada string única, agrupando os gráficos com base em seu nome.
*/
void PlotGraphs() {
    int num_RightNameMacht = 3; // O número de caracteres, a partir da direita, para comparação dos nomes das Worksheets
    
    // Obtenha o nome de todas as WorkSheets no Projeto para comparação...
    out_str("Obtendo o nome de todas as worksheets no Projeto para comparação...");
    vector<string> vector_WksNames;
    foreach(WorksheetPage wksPage in Project.WorksheetPages){
        string name_Page = wksPage.GetName();
        string name_ToMatch = name_Page.Right(num_RightNameMacht);
        vector_WksNames.Add(name_ToMatch);
        //  out_str(name_Page);
    }

    // Itere sobre cada nome
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

## Construção de gráficos a partir de uma lista de nomes
A função abaixo obtém um vetor de nomes de WorksheetsPages, e plota um gráfico agrupando os dados de todas as WorksheetsPages correspondentes.
~~~C
/**
 * Constrói um GraphLayer a partir dos dados das planilhas cujo nomes são passados como argumento
*/
void BuildGraph(vector<string> vector_MatchedNames, string name_Now){
    out_str("Construindo Data Layer para o grupo de: " + name_Now);
    
    // Parâmetros para obtenção dos dados
    string name_DataLayer = "Data";
    int ind_ColX = 0;
    int ind_ColY = 1;

    // Parâmetros para a construção da Legenda
    string str_Exc_Emi = "Exc";
    string str_StartName = "Espectro de Emissão";
    vector temp = {700, 800, 900, 1000, 1100};
    string str_Name = vector_MatchedNames[0];
    string str_AmostName = str_Name.Left(5);
    string str_Legend = "Amostra: " + str_AmostName;
    str_Legend =  str_Legend + "\n" + "\\f:Symbol(l)\\-(" + str_Exc_Emi + ") = " + name_Now + " nm";

    
    GraphPage gp;
    gp.Create("Origin");
    gp.SetLongName(str_StartName + " | " + str_AmostName + " - " + name_Now + "nm");
    GraphLayer gl = gp.Layers();

    double min_X, max_X, min_Y, max_Y;
    bool bool_FirstVal = true;
    for(int j = 0; j < vector_MatchedNames.GetSize(); j++){
        string wksName = vector_MatchedNames[j];
        int int_Ident = atoi(wksName.Left(6).Right(1));
        str_Legend = str_Legend + "\n" + "\\l(" + (j+1) + ") T.T. a " + (int)temp[int_Ident - 1] + " °C";

        // Obtenha a Worksheet com o nome correspondente
        // out_str("Adicionando os dados de: " + wksName + " ao grupo de: " + name_Now);
        WorksheetPage wksPage = Project.WorksheetPages(wksName);
        Worksheet wks = wksPage.Layers(name_DataLayer);
        if(wks){
            DataRange dr;
            // Construa o DataRange e o adicione ao Layer;
            dr.Add(wks, ind_ColX, "X");
            dr.Add(wks, ind_ColY, "Y");
            gl.AddPlot(dr, IDM_PLOT_LINE);

            // Atualize os máximos e mínimos
            Dataset ds_X(wks, ind_ColX);
            Dataset ds_Y(wks, ind_ColY);
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

    FormatGraph(gl, min_X, max_X, min_Y, max_Y, str_Legend);
}
~~~

## Formatação de um gráfico
A função abaixo recebe um objeto GraphLayer e formata este.
~~~C
/**
 * Formata o GraphLayer
*/
void FormatGraph(GraphLayer gl, double min_X, double max_X, double min_Y, double max_Y, string str_Legend){
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

    // Adicione os eixos, e seus ticks, superior e direito ao GraphLayer
    Tree trFormat;
	trFormat.Root.Axes.X.Ticks.TopTicks.show.nVal = 1;
	trFormat.Root.Axes.Y.Ticks.RightTicks.show.nVal = 1;
	int nErr = gl.UpdateThemeIDs( trFormat.Root ) ;
	if(0 != nErr){
		out_str("Fail to Update Theme IDs, theme tree has wrong structure");
        out_str("Verifique a estrutura correta da árvore para este GraphLayer:");
        PlotTree(gl);
    }
	gl.ApplyFormat( trFormat, true, true, true );
    gl.ApplyFormat( trFormat, true, true, true );

    // Adicione uma caixa de texto
    GraphObject go = gl.GraphObjects("Legend");
    go.Text = str_Legend;

	// Acesse o objeto GroupPlot
	GroupPlot gplot = gl.Groups(0);
	if(!gplot){
		out_str("Cannot get the group plot!");
		return;
	}
 
	vector<int> vNester;
 
	// the Nester is an array of types of objects to do nested cycling in the group
	vNester.SetSize(4);  // four types of setting to do nested cycling in the group
	vNester[0] = 0;  // cycling line color in the group
	vNester[1] = 3;  // cycling symbol type in the group
	vNester[2] = 2;  // cycling line style in the group
	vNester[3] = 8;  // cycling symbol interior in the group
 
	gplot.Increment.Nester.nVals = vNester;  // set Nester of the grouped plot
 
	int nNumPlots = gl.DataPlots.Count();
 
	vector<int> vLineColor, vSymbolShape, vLineStyle, vSymbolInterior;
 
	// four vectors for setting line color, symbol shape, line style and symbol interior of grouped plots
	vLineColor.SetSize(nNumPlots);
	vSymbolShape.SetSize(nNumPlots);
	vLineStyle.SetSize(nNumPlots);
	vSymbolInterior.SetSize(nNumPlots);
 
	// setting corresponding values to four vectors
	for(int nPlot=0; nPlot<nNumPlots; nPlot++)
	{
        vSymbolShape[nPlot] = 1;
        vLineStyle[nPlot] = 0;
        vSymbolInterior[nPlot] = 0;
		switch (nPlot)
		{
			case 0:  // setting for plot 0
				vLineColor[nPlot] = SYSCOLOR_BLUE;
				break;
			case 1:  // setting for plot 1
				vLineColor[nPlot] = SYSCOLOR_OLIVE;
				break;
			case 2:  // setting for plot 2
				vLineColor[nPlot] = SYSCOLOR_RED;
				break;
			default:  // setting for other plot
				vLineColor[nPlot] = SYSCOLOR_CYAN;
				break;
		}
	}
 
	trFormat.Root.Increment.LineColor.nVals = vLineColor;  // set line color to theme tree
	trFormat.Root.Increment.Shape.nVals = vSymbolShape;  // set symbol shape to theme tree
	trFormat.Root.Increment.LineStyle.nVals = vLineStyle;  // set line style to theme tree
	trFormat.Root.Increment.SymbolInterior.nVals = vSymbolInterior;  // set symbol interior to theme tree
	trFormat.Root.Line.Width.dVal = 2.0;
 
	if(0 == gplot.UpdateThemeIDs(trFormat.Root) )    
	{
		bool bb = gplot.ApplyFormat(trFormat, true, true);    // apply theme tree
	}else{
		out_str("Fail to Update Theme IDs, theme tree has wrong structure");
        out_str("Verifique a estrutura correta da árvore para este GraphLayer:");
        PlotTree(gl);
    }
}
~~~

## Tratamento dos Erros

~~~C
void PlotTree(GraphLayer gl)
{
    Tree tr;
    tr = gl.GetFormat(FPB_ALL, FOB_ALL, TRUE, TRUE);
    out_str("GraphLayer Tree:\n");
    out_tree(tr,true);
    
    GroupPlot gPlot;
    gPlot= gl.Groups(0); //get first groupplot object
    TreeNode trColormap;
    BOOL bSuccess = gPlot.GetColormap(trColormap);
    out_str("GroupPlot Colormap:\n");
    out_tree(trColormap);
}
~~~
