let FMain = Aliases.TestTD32.FMain;
let PivotGridCust = Aliases.TestTD32.TcxPivotGridCustomizationForm;

// главная процедура
function Main()
{
  let process = RunApplication();
  let WinMain = MaximizeApp(process);
  OpenFormByClick(WinMain, "Форма PivotGrid", 180, 70);
  
  let wnd = FMain.MDIClient.FPivotGrid.pgTest;

  Test1(wnd);
  Test2(wnd);
  Test3();
  Test4(wnd);

  FMain.MDIClient.FPivotGrid.Close(); Delay(2000);
  CloseApp(process);
}

//Выбрать раздел "Форма с PivotGrid"
//В секции Data Fields нажать на правую кнопку мыши
//В меню выбрать пункт 'Show Field List'
//Перенести в окошко поля Square_avg и ShopID
//Закрыть окно 'Field List'
function Test1(wnd)
{
  wnd.ClickR(1,1); 
  wnd.PopupMenu.Click("Show Field List");
  wnd.Drag(518, 17, 1277, 791); // Square_avg -> Field List
  wnd.Drag(610, 9, 1225, 809); // ShopID -> Field List
  PivotGridCust.Close(); // Закрыть окно 'Field List'
  
  Delay(3000);
}
//Перенести в секцию Row Fields поля City и Shop
//Перенести в поля 'Data Fields' поля Income_Sum, Value_custom
//В секции 'Data Fields' нажать на правую кнопку мыши
//В меню выбрать пункт 'Show Field List' 
function Test2(wnd)
{
  wnd.Drag(270, 18, -266, 134); Delay(1000); // City -> Row Fields
  wnd.Drag(486, 15, -454, 128); Delay(1000); // Shop -> Row Fields
  wnd.Drag(478, 13, -469, 2); Delay(1000); // Income_Sum -> Data Fields
  wnd.Drag(508, 6, -402, 13); Delay(1000); // Value_custom -> Data Fields
  
  wnd.ClickR(180, 13);
  wnd.PopupMenu.Click("Show Field List");
  
  Delay(3000);
}
//Выделить поле Square_avg
//В списке выбрать значение 'Data Area'
//Нажать на кнопку 'Add To'
//Выделить поле ShopID
//В списке выбрать значение 'Row Area'
//Нажать на кнопку 'Add To'
function Test3()
{
  let tcxtListBox = PivotGridCust.TcxFieldListListBox;
  let tcxComboBox = PivotGridCust.tcxGroupBox.TcxComboBox;
  let tcxButton   = PivotGridCust.tcxGroupBox.TcxButton;
  
  tcxtListBox.Keys("[Down]"); Delay(500);
  tcxComboBox.ClickItem("Data Area"); Delay(500);
  tcxButton.ClickButton(); Delay(500);
  
  tcxtListBox.Keys("[Down]"); Delay(500);
  tcxComboBox.ClickItem("Row Area"); Delay(500);
  tcxButton.ClickButton(); Delay(500);
  
  Delay(3000);
}

//В поле City нажать на иконку фильтра
//Снять галку с поля (Show All)
//Выбрать город Оренбург
//В секции Row Fields развернуть ветку Оренбург\Магазин 1-3\6
//В секции Column Fields развернуть ветку Россия
//На пересечении строки Оренбург и колонки Россия\Оренбургская обл.\Income_sum ввести значение 300
//На пересечении строки 6 и колонки Россия\Оренбургская обл.\Value_custom ввести значение 600
//На пересечении строки Магазин 2-3 и колонки Россия\Оренбургская обл.\Square_avg ввести значение 400
//Закрыть вкладку "Форма с PivotGrid"
function Test4(wnd)
{
  wnd.Click(120, 88);
  
  let PopupWindow = Aliases.TestTD32.TcxPivotGridFilterPopupWindow
  PopupWindow.TcxPivotGridFilterPopupListBox.TcxCustomInnerCheckListBox.Keys("[Down] [Down][Down][Down] "); 
  PopupWindow.TcxButton.ClickButton();//Выбрать город Оренбург
  Delay(3000);
  
  wnd.Click(10, 111); Delay(1000); //В секции Row Fields развернуть ветку Оренбург\Магазин 1-3\6
  wnd.Click(25, 134); Delay(1000); //В секции Column Fields развернуть ветку Россия
  wnd.Click(564, 36); Delay(1000); 
  
  //На пересечении строки Оренбург и колонки Россия\Оренбургская обл.\Income_sum ввести значение 300
  wnd.Keys("[Down]");
  wnd.Keys("[Down]");
  wnd.Keys("[Right][Right][Right][Right][Right][Right]");
  wnd.Keys("300"); 
  Delay(2000);
  
  //На пересечении строки 6 и колонки Россия\Оренбургская обл.\Value_custom ввести значение 600
  wnd.Keys("[Down]");
  wnd.Keys("[Right]");
  wnd.Keys("600"); 
  Delay(2000);
  
  //На пересечении строки Магазин 2-3 и колонки Россия\Оренбургская обл.\Square_avg ввести значение 400
  wnd.Keys("[Down]");
  wnd.Keys("[Right]");
  wnd.Keys("400"); 
  Delay(2000);

}
  
/////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////
// Запустить приложение
function RunApplication()
{
  return TestedApps.TestTD32.Run();
}

// развернуть приложение на весь экран
function MaximizeApp(p)
{
  w = p.WaitWindow("TFMain", "FMain", -1, 5000);
  
  if (w.Exists)
    {
      w.Activate();
      Log.Picture(w, "MaximizeApp");
    }
    else
      Log.Warning("Incorrect window - MaximizeApp");
      
  w.Maximize();
  return w;
}

//Открыть форму по нажатию клавиши
function OpenFormByClick(w, message, x, y)
{
  if (w.Exists)
    {
      w.Activate();
      w.Click(x, y);
      Log.Picture(w, message);
    }
    else
      Log.Warning("Incorrect window - " + message);   
}

// закрываем приложение
function CloseApp(p)
{  
  p.Close(0);
}