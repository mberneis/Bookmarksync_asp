<!-- #include file = "../inc/asp.inc" -->
<!-- #include file = "../inc/db.inc" -->
<%
restricted ("../bms/frame.asp")
ID=Session("ID")
Session("target") = "bmtarget"
mainmenu = "VIEW"
submenu  = "FRAME"
%>
<script language="javascript">
<!--
var bp = 0;

//if (parent.location.href != window.location.href) {
if (parent.bp == 1) {
	bp = 0;
	parent.location.href='../bms/default.asp';
} 
else {
	bp = 1;
	document.writeln ('<frameset cols="30%,*">');
	document.writeln ('<frame src="../tree/remote.asp" name="bmremote" target="bmtarget">');
	document.writeln ('<frame src="../bms/default.asp" name="bmtarget">');
	document.writeln ('</frameset>');
	document.writeln ('<noframes>' + window.location.href + '</noframes>');
}
//-->
</script>



