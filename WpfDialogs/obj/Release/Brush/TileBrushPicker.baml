   M S B A M L   `   `   `  ���� E  AWpfDialogs, Version=1.1.0.0, Culture=neutral, PublicKeyToken=null'clr-namespace:WpfDialogs
WpfDialogs      WpfDialogs.TileBrushPicker   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     @d2http://schemas.microsoft.com/expression/blend/2008     *localclr-namespace:WpfDialogs     -   $��TbPkr�����#EEE  � 	Resources   �� MergedDictionaries �� Source �� 
TargetType �� Property �� Value��  �� $ ..\ExpanderStyle.xaml?�%G   '��       �� " ��.���  � ��$ Center=���  � ��$ 18(�.���� $��1��$��Disabled=�$
��Auto=�.���� .+�� .+�9� $��Hatch Shape�� 9� 
IsExpanded . .���� -    ��Name$ LstTile��$��5q�$	��120��$��Disabled=����� .R�� 
������� � Orientation$ 
Horizontal=�	 � ActualWidth#��z	 
 K� IsItemsHost
 . ���� .R���  ��Name$
 Brdr��
���  ��$��49��$��49��$��5q�._��� 
���   �$	��0.5q�$��3q�����  � ToolTip ��  �� Path$% (local:HatchBrushes.HatchStyle)� �� Triggers ��  �� Value$
 True�� �� Binding ��  �� RelativeSource �� *TemplatedParent+$ 
IsSelected�.a���  �� 
TargetName$
 Brdr�� � ��$ 3q�   WpfDialogs.TransformBox  -     Name$ ShapeTransPkr��$��22��$��22��$ Shape Transformations��$��Right=�$	��Top=�
���   �$��1q�$��	0 13 13 0q�{��� *	IsEnabled+ �� ElementName$ StkPnlTileOptions���� -   $��StkPnlTileOptions��.+�9� $��Hatch Parameters��.���  � ColumnDefinitions �� $��95 ��� �� $��95 ���  � RowDefinitions �� �� .+��� $��
Background��   WpfDialogs.ColorBox  -     Name$ PkrBackground��$��1��
���   �$��1q�   Brush
 �  ��$��5q����� *ActualHeight+$ UdPenThickness���� $��
Foreground��$��2��  -   $ PkrForeground��$��3��
���   �$��1q�
 � ���$��5q����� *ActualHeight+$ UdPenThickness���� $��	Thickness��$ Pen Thickness��$��1��   WpfDialogs.DoubleUpDown  -     Name$ UdPenThickness��$ Pen Thickness��   WpfDialogs.UpDownBase   Value$	 1.0X�$��1��$��1��$��5q�   Minimum$ 0Q��� $��	Transform��$��1��$��2��  -   $ HatchTransPkr��$��1��$��3��
���   �$��1q�$��5q�9� $��Tile Options��.���  �� $��50 ��� �� $��50 ���  �� �� .+��� $
��Mode���� -	    ��Name$ cmbTileMode��$��1��$
��12,5q��� $��Stretch��$��2���� -
   $ 
cmbStretch��$��3��$
��12,5q��� $��AlignX��$��1���� -   $ 	cmbAlignX��$��1��$��1��$
��12,5q��� $��AlignY��$��1��$��2���� -   $ 	cmbAlignY��$��1��$��3��$
��12,5q�9� $��	View Port��.���  �� $��50 ��� �� $��50 ���  �� �� .+��� $
��Left��  -   $ UdViewPortX��$��1��$
��12,5q��� $	��Top��$��2��  -   $ UdViewPortY��$��3��$
��12,5q��� $��Width��$��1��  -   $ UdViewPortWidth��$��1��$��1��$
��12,5q�$ 0Q��� $��Height��$��1��$��2��  -   $ UdViewPortHeight��$��1��$��3��$
��12,5q�$ 0Q�9� $��View Box��.���  �� $��50 ��� �� $��50 ���  �� �� .+��� $
��Left��  -   $ 
UdViewBoxX��$��1��$
��12,5q��� $	��Top��$��2��  -   $ 
UdViewBoxY��$��3��$
��12,5q��� $��Width��$��1��  -   $ UdViewBoxWidth��$��1��$��1��$
��12,5q�$ 0Q��� $��Height��$��1��$��2��  -   $ UdViewBoxHeight��$��1��$��3��$
��12,5q�$ 0Q�    WpfDialogs.TransformPicker  -     Name$ TransformPicker��$	��0,5q�   	Transform �� $ Brush.RelativeTransform�$ TbPkr�� �� Mode$ TwoWay=�