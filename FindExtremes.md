~~~C

void FindExtreme(){
  // Parâmetros para obtenção dos dados
  string name_DataLayer = "Data";
  int ind_ColDataX = 3, ind_ColDataY = 4;

  // Parâmetros para identificação de extremos
  double trheshold_FirstDer = 6;            // Módulo máximo para a primeira derivada de um extremo
  double trheshold_SecDer = 200;            // Limite para a segunda derivada

  // Parâmetros para identificar absorção de um extremo pelo outro
  double trheshold_YExtr = 0.01;            // Módulo da variação mínima de intensidade entre dois extremos;
  double threshold_D = 0.02;                // Distância mínima entre dois extremos

  foreach(WorksheetPage wksPage in Project.WorksheetPages){
    // Obtenha os dados da Worksheet
    Worksheet wks = wksPage.Layers(name_DataLayer);
    out_str("\n\nAcessando WorksheetPage com Long Name: " + wksPage.GetLongName());
    out_str("Acessando coluna com Long Name: " + wks.Columns(ind_ColDataY).GetLongName() + "\n");
    Dataset dsY(wks, ind_ColDataY);
    Dataset dsX(wks, ind_ColDataX);
    int size = dsY.GetSize();
    double y_LastExtreme = dsY[0];
    double x_LastExtreme = dsX[0];

    /* Matrizes para coleta dos extremos
    *   {Tipo, Valor, Posição, Distância entre até o extremo anterior ,Primeira Derivada, Segunda Derivada}
    *   Tipo = 0: Mínimo Local
    *   Tipo = 1: Máximo Local
    */
    int num_Cols = 6;
    matrix mat_Extr(1,num_Cols);
    int l = 0;
    int tipo;
    bool isMax = true;
    int dif = -1;
    for(int i = 1; i < size - 1; i++){
      // Obtenha o ponto e seus primeiros vizinhos
      double y_Prev = dsY[i-1], y_Curr = dsY[i], y_Next = dsY[i+1];
      double x_Prev = dsX[i-1], x_Curr = dsX[i];

      // Verifique a presença de um extremo
      double increment = x_Curr - x_Prev;
      double firstDerivative = (y_Next - y_Prev)/(2*increment);
      if(abs(firstDerivative) < trheshold_FirstDer){
        // Obtenha as características no ponto
        double del_YExtr = y_Curr - y_LastExtreme;
        double dist_Extr = (del_YExtr^2 + (x_Curr - x_LastExtreme)^2)^0.5;
        double secondDerivative = (y_Prev - 2*y_Curr + y_Next)/increment^2;
        
        bool isExtreme = dist_Extr > threshold_D;
        isExtreme = isExtreme && (abs(del_YExtr) > trheshold_YExtr);
        isExtreme = isExtreme && (abs(secondDerivative) > trheshold_SecDer);
        // Atualize os valores no vetor
        bool rc = true;
        bool subPoint = false;
        if(dif*y_Curr > dif*y_LastExtreme){
          subPoint = true;
        }else if(isExtreme){
          dif = dif*(-1);
          isMax = !isMax;
        }
        
        if(subPoint || isExtreme){
          if(!subPoint){
            l++;
            rc = mat_Extr.SetSize(l + 1,num_Cols,true);
            if(!rc){
              out_str("Falha em redimensionar a matriz");
            }
          }
          y_LastExtreme = y_Curr;
          x_LastExtreme = x_Curr;
          mat_Extr[l][0] = dif;
          mat_Extr[l][1] = y_Curr;
          mat_Extr[l][2] = x_Curr;
          mat_Extr[l][3] = dist_Extr;
          mat_Extr[l][4] = firstDerivative;
          mat_Extr[l][5] = secondDerivative;
        }
      }
    }

    // Imprima os extremos encontrados
    for(i = 0; i < l; i++){
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
    }
  }
}

~~~
