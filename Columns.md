~~~C

void SetColumnValue(){
  // Variaveis para a Worksheet processada
  string name_DataLayer = "Data";

  // Variáveis para a Coluna criada
  string str_form = "1240,86/col(A)";
  string str_ShortName = "";
  string str_LongName = "";
  string str_Units = "";
  string str_Coment = "";

  foreach(WorksheetPage wksPage in Project.WorksheetPages){
    out_str("Acessando WorksheetPage com Long Name: " + wksPage.GetLongName());
    
    // Obtenha a Worksheet
    Worksheet wks = wksPage.Layers(name_DataLayer);
    if(wks){
      // Adicione a coluna
      int int_IndCol = wks.AddCol(str_ShortName);
      Column col = wks.Columns(int_IndCol);

      // Adicione propriedades à coluna
      col.SetType(OKDATAOBJ_DESIGNATION_X);
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
