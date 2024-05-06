~~~C

void FindExtreme(){
  // Parâmetros para obtenção dos dados
  string name_DataLayer = "Data";
  int ind_ColDataX = 3;
  int ind_ColDataY = 4;

  // Parâmetros para identificação de extremos
  double trheshold_FirstDer = 6;            // Módulo máximo para a primeira derivada de um extremo
  double trheshold_SecDer = 200;            // Limite para a segunda derivada

  // Parâmetros para identificar absorção de um extremo pelo outro
  double trheshold_YExtr = 0.01;            // Módulo da variação mínima de intensidade entre dois extremos;
  double threshold_D = 0.02;                // Distância máxima entre dois extremos
  double threshold_YMax = 0.02;             // Valor mínimo para a consideração de um máximo

  foreach(WorksheetPage wksPage in Project.WorksheetPages){
    // Obtenha os dados
    Worksheet wks = wksPage.Layers(name_DataLayer);
    out_str("\n\nAcessando WorksheetPage com Long Name: " + wksPage.GetLongName());
    out_str("Acessando coluna com Long Name: " + wks.Columns(ind_ColDataY).GetLongName() + "\n");
    Dataset dsY(wks, ind_ColDataY);
    Dataset dsX(wks, ind_ColDataX);
    int size = dsY.GetSize();
    double lastExtreme = dsY[0];
    double pos_lastExtreme = dsX[0];
    bool isMax = true;

    for(int i = 1; i < size - 1; i++){
      double prev = dsY[i-1];
      double curr = dsY[i];
      double next = dsY[i+1];
      double increment = dsX[i] - dsX[i-1];
      double firstDerivative = (dsY[i+1] - dsY[i-1])/(2*increment);           // Aproximação para a primeira derivada
      double secondDerivative = (dsY[i-1] - 2*curr + dsY[i+1])/increment^2;   // Aproximação para a segunda derivada

      bool teste_Y = abs(curr - lastExtreme) > trheshold_YExtr;
      bool teste_D = ((curr - lastExtreme)^2 + (dsX[i] - pos_lastExtreme)^2)^0.5 > threshold_D;
      bool test_FirstDer = abs(firstDerivative) < trheshold_FirstDer;
      bool test_SecDer = abs(secondDerivative) > trheshold_SecDer;
      if(teste_D && test_FirstDer && test_SecDer){
        bool extremeFinded = false;
        
        if(teste_Y){
          if(isMax){
            if(curr > prev && curr > next){
              extremeFinded = true;
            }
          }else{
            if(curr < prev && curr < next){
              extremeFinded = true;
            }
          }
        }
        if(curr < prev && curr < next && curr < lastExtreme){
          extremeFinded = true;
          isMax = false;
        }
        if(curr > prev && curr > next && curr > lastExtreme){
          extremeFinded = true;
          isMax = true;
        }

        if(extremeFinded){
          bool test_YMax = isMax && curr < threshold_YMax;
          if(test_YMax){
            // Absorva o ponto.
          }else{
            if(isMax){
              out_str("Máximo Local:")
            }else{
              out_str("Mínimo Local:")
            }
            isMax = !isMax;
            pos_lastExtreme = dsX[i];
            lastExtreme = curr;
            out_str("\tValor: " + curr);
            out_str("\tNa posição: " + dsX[i]);
            out_str("\tValor da Primeira Derivada: " + firstDerivative);
            out_str("\tValor da Segunda Derivada: " + secondDerivative);
          }
        }
      }
    }
  }
}

~~~