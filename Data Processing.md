# Processamento de Dados em Massa
A seguinte função adiciona uma coluna de dados a todas as WorksheetsPages no arquivo Origin, adicionando algumas características a essas colunas.
~~~C
void CreateColumn(string name_DataLayer, string str_form, string str_ShortName, string str_LongName, string str_Units, string str_Coment, int ColType){
  // 0: Y
  // 1: None
  // 2: Y Error
  // 3: X
  // 4: L
  // 5: Z
  // 6: X Error

  foreach(WorksheetPage wksPage in Project.WorksheetPages){
    out_str("Acessando WorksheetPage com Long Name: " + wksPage.GetLongName());
    
    // Obtenha a Worksheet
    Worksheet wks = wksPage.Layers(name_DataLayer);
    if(wks){
      // Adicione a coluna
      int int_IndCol = wks.AddCol(str_ShortName);
      Column col = wks.Columns(int_IndCol);

      // Adicione propriedades à coluna
      col.SetType(ColType);
      col.SetFormula(str_form, AU_AUTO);
      col.ExecuteFormula();
      col.SetLongName(str_LongName);
      col.SetUnits(str_Units);
      col.SetComments(str_Coment);
    }else{
      out_str("A planilha " + wks.GetLongName() + " não possui uma Worksheet com o nome " + name_DataLayer);
    }
  }
}
~~~

## Visualizando extremos
A seguinte função encontra os pontos de extremos (máximos e mínimos), plota os gráficos dos dados e adiciona marcadores de dados em cada ponto de mínimo encontrado. Para tanto a função calcula uma aproximação para as derivadas da função, até da terceira ordem, como também calcula média móvel dos dados para amenizar o ruído dos dados. 
~~~C
void ExtremesVisualization(){
  // Parâmetros para obtenção dos dados
  string name_DataLayer = "Data";
  int ind_ColDataX = 3, ind_ColDataY = 4;


  foreach(WorksheetPage wksPage in Project.WorksheetPages){
    // Obtenha os dados da Worksheet
    Worksheet wks = wksPage.Layers(name_DataLayer);
    out_str("\n\nAcessando WorksheetPage com Long Name: " + wksPage.GetLongName());
    out_str("Acessando coluna com Long Name: " + wks.Columns(ind_ColDataY).GetLongName() + "\n");
    Dataset dsY(wks, ind_ColDataY);
    Dataset dsX(wks, ind_ColDataX);

    // Obtenha a média móvel dos valores
    vector vec_Y;
    vec_Y.Append(MovingAvg(dsY, 5));
    vector vec_X(dsX);

    // Obtenha a matriz com extremos
    matrix mat_Extr = FindExtreme(vec_X, vec_Y);
    int l = mat_Extr.GetNumRows();

    // Plote os dados obtidos
    GraphPage gp;
    gp.Create("Origin");
    gp.SetLongName("Emissão vs Energia | " + wksPage.GetLongName() + "nm");
    GraphLayer gl = gp.Layers();
    DataRange dr;
    dr.Add(wks, ind_ColDataX, "X");
    dr.Add(wks, ind_ColDataY, "Y");
    gl.AddPlot(dr, IDM_PLOT_LINE);
    gl.Rescale();

    // Itere sobre cada extremo
    DataPlot dp = gl.DataPlots();
    vector<int> vnBegin;
    vector<int> vnEnd;
    BOOL bInvalidate = TRUE;
    BOOL bUndo = TRUE;
    int num_Mins = 0;
    for(int i = 0; i < l; i++){
      int tipo = mat_Extr[i][0];
      double val = mat_Extr[i][1];
      if(tipo == -1){
        out_str("Mínimo Local:");
        num_Mins++;
        if(num_Mins == 1){
          vnBegin.Add(val);
        }else{
          vnBegin.Add(val);
          vnEnd.Add(val);
        }
      }else{
        out_str("Máximo Local:");
      }
      out_str("\tValor: " + mat_Extr[i][2]);
      out_str("\tNa posição: " + mat_Extr[i][3]);
      out_str("\tDistância até o extremo anterior: " + mat_Extr[i][4]);
      out_str("\tPrimeira Derivada: " + mat_Extr[i][5]);
      out_str("\tSegunda Derivada: " + mat_Extr[i][6]);
      out_str("\tTerceira Derivada: " + mat_Extr[i][7]);
      out_str("\tPrimeira Derivada Ajustada: " + mat_Extr[i][8]);
    }
    int nRet = dp.AddDataMarkers(vnBegin, vnEnd, bInvalidate, bUndo);
    ASSERT(0 == nRet);
  }
}
~~~
### Obtendo a média móvel de um conjunto de dados
~~~C
vector MovingAvg(Dataset dsY, int n){
  vector vec_Y;
  int l = n - 1;
  for(int i = 0; i < dsY.GetSize(); i++){
    if(i < n) l = i;
    double avg = 0;
    for(int j = i - l; j <= i; j++){
      avg = avg + dsY[j];
    }
    vec_Y.Add(avg/(l + 1));
  }
  return vec_Y;
}
~~~
### Encontrando extremos de um conjunto de dados
A seguinte função encontra os extremos de um conjunto de dados passado através de um vetor. Para tanto a função usa de alguns limites arbitrários (que dependem de cada tipo de dado) para julgar se um ponto é um limite significativo ou não.
~~~C
matrix FindExtreme(vector dsX, vector dsY){
  // Parâmetros para identificação de extremos
  double trheshold_FirstDer = 6;            // Módulo máximo para a primeira derivada de um extremo
  double trheshold_SecDer = 200;            // Limite para a segunda derivada

  // Parâmetros para identificar absorção de um extremo pelo outro
  double trheshold_YExtr = 0.01;            // Módulo da variação mínima de intensidade entre dois extremos;
  double threshold_D = 0.02;                // Distância mínima entre dois extremos
  
  int size = dsY.GetSize();
  double y_LastExtreme = dsY[0],  x_LastExtreme = dsX[0];
  double y_LastPoint = dsY[0], x_LastPoint = dsX[0];

  /* Matrizes para coleta dos extremos
  *   {Tipo, Número da iteração, Valor, Posição, Distância entre até o extremo anterior ,Primeira Derivada, Segunda Derivada, Terceira Derivada, Primeira Derivada Corrigida}
  *   Tipo = 0: Mínimo Local
  *   Tipo = 1: Máximo Local
  */
  int num_Cols = 9;
  matrix mat_Extr(1,num_Cols);
  int l = 0;
  int tipo;
  bool isMax = true;
  int dif = -1;
  for(int i = 2; i < size - 2; i++){
    // Obtenha o ponto e seus primeiros vizinhos
    double y_Prev2 = dsY[i-2], y_Prev1 = dsY[i-1], y_Curr = dsY[i], y_Next1 = dsY[i+1], y_Next2 = dsY[i-2];
    double x_Prev = dsX[i-1], x_Curr = dsX[i];

    // Verifique a presença de um extremo
    double increment = x_Curr - x_Prev;
    double firstDerivative = Derivatives(1,y_Next1, y_Prev1, increment, y_Curr);
    if(abs(firstDerivative) < trheshold_FirstDer){
      // Obtenha as características no ponto
      double secondDerivativePrev = Derivatives(2,y_Curr,y_Prev2,increment,y_Prev1);
      double secondDerivative = Derivatives(2,y_Next1, y_Prev1, increment,y_Curr);
      double secondDerivativeNext = Derivatives(2,y_Next2,y_Curr,increment,y_Next1);
      double thirdDerivative = Derivatives(1,secondDerivativeNext,secondDerivativePrev,increment,secondDerivative);
      double adjFirstDeriv = firstDerivative - (increment^2)*(thirdDerivative)/6;
      double del_YExtr = y_Curr - y_LastPoint;
      double dist_Extr = (del_YExtr^2 + (x_Curr - x_LastPoint)^2)^0.5;
      bool isExtreme = dist_Extr > threshold_D;
      isExtreme = isExtreme && (abs(del_YExtr) > trheshold_YExtr);
      isExtreme = isExtreme && (abs(secondDerivative) > trheshold_SecDer);

      // Identifique o tipo de ponto
      bool subPoint = false;
      if(dif*y_Curr > dif*y_LastPoint){
        subPoint = true;
        dist_Extr = ((y_Curr - y_LastExtreme)^2 + (x_Curr - x_LastExtreme)^2)^0.5;
      }else if(isExtreme){
        dif = dif*(-1);
        isMax = !isMax;
        y_LastExtreme = y_Curr;
        x_LastExtreme = x_Curr;
      }
      
      if(subPoint || isExtreme){
        if(!subPoint){
          l++;
          bool rc = mat_Extr.SetSize(l + 1,num_Cols,true);
          if(!rc){
            out_str("Falha em redimensionar a matriz");
          }
        }
        y_LastPoint = y_Curr;
        x_LastPoint = x_Curr;
        mat_Extr[l][0] = dif;
        mat_Extr[l][1] = i - 1;
        mat_Extr[l][2] = y_Curr;
        mat_Extr[l][3] = x_Curr;
        mat_Extr[l][4] = dist_Extr;
        mat_Extr[l][5] = firstDerivative;
        mat_Extr[l][6] = secondDerivative;
        mat_Extr[l][7] = thirdDerivative;
        mat_Extr[l][8] = adjFirstDeriv;
      }
    }
  }
  return mat_Extr;
}
~~~
### Aproximação para a primeira e segunda derivada
~~~C
double Derivatives(int n , double y_Next, double y_Prev, double increment, double y_Curr){
  if(n == 1){
    return  (y_Next - y_Prev)/(2*increment);
  }else if (n == 2){
    return (y_Prev - 2*y_Curr + y_Next)/(increment^2);
  }else{
    return "A ordem da derivada deve ser igual a 1 ou 2.";
  }
}
~~~
