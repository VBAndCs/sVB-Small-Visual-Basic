   M S B A M L   `   `   `  ���� E  AWpfDialogs, Version=1.1.0.0, Culture=neutral, PublicKeyToken=null'clr-namespace:WpfDialogs
WpfDialogs      WpfDialogs.FontChooser   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     *localclr-namespace:WpfDialogs     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     @d2http://schemas.microsoft.com/expression/blend/2008     .��� -   $��12q�  � RowDefinitions  �� $
��Auto ��� $��5* ��� $��2* ��� $
��Auto �.+�� $��0��$��0�� � ColumnDefinitions �� $
��Auto ��� $��2* ��� $��12 ��� $
��Auto ��� $��* �  �� $��30 ��� $��5 ��� $��30 ��� $��5 ��� $��30 ��� $��5 �.+��� $��_Font: ���� -   $��fontFamilyList��$��1��#��� �$��Stretch=� �� 
IsEditable .  �� StaysOpenOnEdit . �� $��2��$��
_Typeface:���� -   $��typefaceList��$��2��$��1��#��� �$��Stretch=� . �� $��4��$��_Size:���� -   $��sizeList��$��4��$��1��#��� �$��Stretch=� .  . ]� $��4��$��5��.+��� -   $��underlineCheckBox��$	��Top=�$��	Underline���� -   $��baselineCheckBox��$	��Top=�$��Baseline���� -   $��strikethroughCheckBox��$	��Top=�$��Strikethrough���� -   $��overlineCheckBox��$��Overline���� -	   $��
tabControl��$��0��$��1��$��0,12,0,0q�.@��� -
   $��typographyTab��$��Typographic Features��.���  �� $
��Auto ��� $��8 ���   �� $
��Auto ��� .+��� $��0��$��0��$��Feature���� -   $��featureList��$��0��$��1��$	��230��� $��2��$��0�� �� $��1 ��� $��96 ��� .+��� $��1��$��	Selection���� $��2��$��Sample���� -    ��Name$ TypographyBorder��$��2��$��1��$��1q�#��Z�._�� }�. .+��� -    ��Name$ 
ScvFeature��$
��Auto=�.��� -   $��TypographyFeatureGrid�� �� $	��200 ��� �� -   $ SelectionBoreder��$��2��$��1��$��1q�
���   �$	��Top=�$
��Left=� J� IsHitTestVisible .  
����� � J� Opacity$	 0.1Q��� -   $��
samplesTab��$!��Family and typeface samples��.���  �� �� $��12 ���   �� $
��Auto ��� .+��� $��0��$��0���� .�FontFamily: �� -    �Name$ fontFamilyNameRun���� -   $��fontFamilySamples��$��0��$��1��	 �� 
IsReadOnly	 . "
 �� VerticalScrollBarVisibility$

 Auto=��� $��2��$��0���� .�	Typeface: �� -   $ typefaceNameRun��.�Points�� -   $��typefaceSamples��$��2��$��1��	 . $

 Auto=��� -   $��descriptiveTextTab��$��DescriptiveText��.���  �� $��3* ��� $��8 ��� $��2* �  �� $
��Auto ��� .+��� $��0��$��0��$��Font description:���� -   $��fontDescriptionBox��$��0��$��1��	 . $

 Auto=��� $��2��$��0��$��License:���� -   $��fontLicenseBox��$��2��$��1��	 .  �� TextWrapping$ WrapWithOverflow=�$

 Auto=�]� $��0��$��2��$��0,12,0,0q�.+��� -   $��previewTextBox��$��16(� �� AcceptsReturn .  �� TextAlignment$ Center=�$
 Wrap=�$

 Auto=��� $��0��$��3�� �� Orientation$ 
Horizontal=�$��Right=�$��0,12,0,0q�.+��� -   $��OK��$��86�� �� 	IsDefault . �� -   $��Cancel��$��86��