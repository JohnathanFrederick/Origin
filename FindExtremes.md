~~~C


void FindExtreme(){
  string name_DataLayer = "Data";
  int ind_ColDataX = 3;
  int ind_ColDataY = 4;
  double threshhold = 0.01;

  foreach(WorksheetPage wksPage in Project.WorksheetPages){
    out_str("\nAcessando WorksheetPage com Long Name: " + wksPage.GetLongName());
    Worksheet wks = wksPage.Layers(name_DataLayer);
    out_str("\nAcessando coluna com Long Name: " + wks.Columns(ind_ColDataY).GetLongName() + "\n");
    Dataset dsY(wks, ind_ColDataY);
    Dataset dsX(wks, ind_ColDataX);
    int size = dsY.GetSize();
    double lastExtreme = dsY[0];
    bool isMax = true;

    for(int i = 1; i < size -1; i++){
      double prev = dsY[i-1];
      double curr = dsY[i];
      double next = dsY[i+1];

      if(abs(curr - lastExtreme) > threshhold){
        if(curr > prev && curr > next && isMax){
          out_str("Máximo Local: " + curr + " na posição: " + dsX[i]);
          lastExtreme = curr;
          isMax = false;
        }else if(curr < prev && curr < next && !isMax){
          out_str("Mínimo Local: " + curr + " na posição: " + dsX[i]);
          lastExtreme = curr;
          isMax = true;
        }
      }
    }
  }
}

~~~
