~~~C

void Teste(){
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

    vector mat_Y;
    mat_Y.Append(MovingAvg(dsY, 5));
    vector mat_X(dsX);

    matrix mat_Extr = FindExtreme(mat_X, mat_Y);
    int l = mat_Extr.GetNumRows();

    // Imprima os extremos encontrados
    for(int i = 0; i < l; i++){
      int tipo = mat_Extr[i][0];
      if(tipo == -1){
        out_str("Mínimo Local:");
      }else{
        out_str("Máximo Local:");
      }
      out_str("\tValor: " + mat_Extr[i][1]);
      out_str("\tNa posição: " + mat_Extr[i][2]);
      out_str("\tDistância até o extremo anterior: " + mat_Extr[i][3]);
      out_str("\tPrimeira Derivada: " + mat_Extr[i][4]);
      out_str("\tSegunda Derivada: " + mat_Extr[i][5]);
      out_str("\tTerceira Derivada: " + mat_Extr[i][6]);
      out_str("\tPrimeira Derivada Ajustada: " + mat_Extr[i][7]);
    }
  }
}

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
  *   {Tipo, Valor, Posição, Distância entre até o extremo anterior ,Primeira Derivada, Segunda Derivada, Terceira Derivada, Primeira Derivada Corrigida}
  *   Tipo = 0: Mínimo Local
  *   Tipo = 1: Máximo Local
  */
  int num_Cols = 8;
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
        mat_Extr[l][1] = y_Curr;
        mat_Extr[l][2] = x_Curr;
        mat_Extr[l][3] = dist_Extr;
        mat_Extr[l][4] = firstDerivative;
        mat_Extr[l][5] = secondDerivative;
        mat_Extr[l][6] = thirdDerivative;
        mat_Extr[l][7] = adjFirstDeriv;
      }
    }
  }
  return mat_Extr;
}

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
