   M S B A M L   `   `   `  ���� F  BExcelToJson, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null)clr-namespace:ExcelToJsonExcelToJson      ExcelToJson.MainWindow   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     +localclr-namespace:ExcelToJson       � Title$  ExcelToJson��$	��350��$	��620��
�������$	��620��$	��350��.��� $��0,0,2,0q� � ColumnDefinitions �� .+��� -   $��选择��$
��Left=�$��
503,26,0,0q�$	��Top=�$��82��$��23���� -    ��Name$ SrcPath��$
��Left=�$��23��$��	94,26,0,0q� �� TextWrapping$
 Wrap=�$	��Top=�$	��404���� -    ��Name$ SrcText��$
��Left=�$��	21,27,0,0q�$
��Wrap=�$��源文件夹：��$	��Top=�$��23��{�.   �� TextAlignment$ Center=� � Cursor$
 None���� -   $��选择��$
��Left=�$��
503,54,0,0q�$	��Top=�$��82��$��23���� -   $ 
TargetPath��$
��Left=�$��23��$��	94,54,0,0q�$
 Wrap=�$	��Top=�$	��404���� -   $ 
TargetText��$
��Left=�$��	21,54,0,0q�$
��Wrap=�$��目标文件夹：��$	��Top=�$��23��{�.  $ Center=�$
 None���� -    ��Name$ LogList��$
��Left=�$	��165��$��	21,94,0,0q�$	��Top=�$	��477�����	#FFE2E1E1 �� View �� .P��� $��Excel文件��	 �� Width$		 467���� -   
 ��Name$
 ConfirmButton��$��一键转换��$
��Left=�$��505,192,0,0q�$	��Top=�$��80��$��67��$��14(��� -	    ��Name$ IsAllWorkBookCheckBox��$��所有工作簿
��$
��Left=�$��
505,94,0,0q�$	��Top=� d� 	IsChecked$
 TrueY�$��24���� -
   $ LogText��$
��Left=�$��33��$��
21,276,0,0q�$
 Wrap=�$	��Top=�$	��557��
���  �� �� SelectionBrush
 �����
�������{�.  �� -    ��Name$ TransformTypeCombo��$
��Left=�$��505,123,0,0q�$	��Top=�$��80�� �� SelectedIndex$ 0��$��24��.@��� $��转为数组���� $��转为对象���� -   $ CreatCPPCode��$��	生成CPP��$
��Left=�$��505,152,0,0q�$	��Top=�$��80��$
 TrueY��� -   $ CreatCSharpCode��$��生成C#��$
��Left=�$��505,172,0,0q�$	��Top=�$��80��$
 TrueY�