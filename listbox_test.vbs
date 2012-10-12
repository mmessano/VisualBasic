<SCRIPT LANGUAGE="VBSCRIPT">
<!--
	Sub window_onload
		Form1.elements(0).focus
	End Sub

	Sub Button1_onclick
		MsgBox "The " & Form1.Select1.options(Form1.Select1.selectedIndex).value & " is a fine automobile."
	End Sub
-->
</SCRIPT>

<FORM NAME="Form1">Please select a make:
	<BR>
	<SELECT NAME="Select1" SIZE=10 VALUE="Selection">
		<OPTION VALUE="Acura" SELECTED>Acura
		<OPTION VALUE="Audi"> Audi
		<OPTION VALUE="BMW">BMW
		<OPTION VALUE="Buick">Buick
		<OPTION VALUE="Cadillac">Cadillac
		<OPTION VALUE="Chevrolet">Chevrolet
		<OPTION VALUE="Chrysler">Chrysler
		<OPTION VALUE="Dodge">Dodge
		<OPTION VALUE="Ford">Ford
		<OPTION VALUE="Honda">Honda
	</SELECT>
	<P><INPUT TYPE=BUTTON NAME="Button1" VALUE="Continue">
</FORM>