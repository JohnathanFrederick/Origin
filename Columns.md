~~~C

void CreateColumns(){
  CreateColumn("Data", "1240,86/col(A)", "", "Energia", "eV", "Energia associada ao comprimento de onda", 3);
  CreateColumn("Data", "col(B)", "", "Intensidade Relativa", "", "Intensidade Relativa da Emissão", 0);
}

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
