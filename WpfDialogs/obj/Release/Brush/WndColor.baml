   M S B A M L   `   `   `  ���� E  AWpfDialogs, Version=1.1.0.0, Culture=neutral, PublicKeyToken=null'clr-namespace:WpfDialogs
WpfDialogs     WpfDialogs.WndColor   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     *localclr-namespace:WpfDialogs       � Title$  Color Dialog�� � SizeToContent$ WidthAndHeight=� � 
ResizeMode$ NoResize=� � WindowStartupLocation$ CenterScreen=�.���  � ColumnDefinitions �� $	��160 ��� $	��450 ���  � RowDefinitions �� $	��430 ��� $��70 �.+�� $��2��$��4q� �� $
��Auto ��� .+���  �� TextAlignment$ Center=�$��Recent Brushes��$��2,1,2,4q��� -    ��Name$ 
lstBrushes��$��1��
���   �$��1q�����  �� 
TargetType" ��.��� 	 �� Property	 � ��
 �� Value
 ��  �� 
TargetType" ��.R���  ��Name$ Border��$��5,5,5,0q�$	��120��$	��120��
������ ._��� -   $��2q�����  �� Triggers P�  P� Property �@J�IsMouseOver P� Value . . ���  �� 
TargetName$ Border��	 � ��

 �   ��� $ Border��	 � ��$
 1q�P�  �@��
IsSelected . . ��� $ Border��	 � ��

 �   ��� $ Border��	 � ��$
 2q��� $��1��$��2��.@��� $��Preview��.���� 
���   �$	��1.5q�._�� .+��� -    ��Name$ lblBrushPreview���� -    ��Name$ BtnReset��
���  ��
�������$��Right=�$	��Top=�$��5q�$��0q�$��25��$��30�� � ToolTip$ Remove Brush��.��X�� $��	Xaml Code��.���� -    ��Name$ txtXaml��$��24(� �� 
IsReadOnly . " �� VerticalScrollBarVisibility$
 Auto=�$ �� HorizontalScrollBarVisibility$
 Auto=��� -    ��Name$ TabControl1��$��2��$	��350��$	��5,0q�.@��� $��Color Brush��.��  WpfDialogs.ColorPicker  -     Name$	 Pkr��$	��350��   Brush ��  �� ElementName$ lblBrushPreview�� �� Path$ 
Background� �� Mode$ OneWayToSource=��� $��
Tile Brush��.��   WpfDialogs.TileBrushPicker  -     Name$ 
TilePicker��   Brush �� $ lblBrushPreview��$ 
Background�$ OneWayToSource=�    
PreviewBox  �� $ lblBrushPreview���� $��Image Brush��.��!  WpfDialogs.ImageBrushPicker  -	   !  Name$! ImagePicker��"   Brush" �� $ lblBrushPreview��$ 
Background�$ OneWayToSource=�#   
PreviewBox# �� $ lblBrushPreview���� $��1��$��2��
���   �$��1q�$��5q�._��� $��5q�$ �� Orientation$$ 
Horizontal=�$��Center=�.+��� -
   $ BtnOk��$	��120��$
��10,5q�% �� 	IsDefault% . .��OK�� -   $ 	BtnCancel��$	��120��$
��10,5q�& �� IsCancel& . .��Cancel