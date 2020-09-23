<div align="center">

## A Basic Shopping Cart


</div>

### Description

This code is just to show you how you can make a simple quick and easy shopping cart for your site. Teaches you how to use session variables and the dictionary object (no, its not for spell checking either!)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dustin R Davis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dustin-r-davis.md)
**Level**          |Intermediate
**User Rating**    |4.3 (51 globes from 12 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dustin-r-davis-a-basic-shopping-cart__4-7538/archive/master.zip)

### API Declarations

Please do not steal code!


### Source Code

```
<%
'''''''''''''''''''''''''''''''''''''''''''''''''
' A Simple Shopping Cart						'
' Coded By: Dustin Davis						'
' Date: 05/11/2002								'
'												'
'This is just a simple example of how to start	'
'Shopping cart for your site. You can Add/Delete'
'and View your items							'
'												'
'this also shows you how to use Session			'
'variables and the dictionary object			'
'Please do not steal code, give credit where it	'
'is do!											'
'''''''''''''''''''''''''''''''''''''''''''''''''
Dim Basket  ' This will hold our shopping cart information
dim tmpItems ' This will hold all of the items we have in our Shopping Cart
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This If statement will check to see if anything is in the querystring
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request.QueryString("action") = "" then
	'Write a form to take in information
	Response.Write "<HTML><BODY><FORM NAME='addit' ACTION='./Basket.asp?action=Add' METHOD='POST'>"
	Response.Write "<INPUT TYPE='TEXT' NAME='ITEM' VALUE=''><INPUT TYPE='SUBMIT' VALUE='ADD'></FORM>"
	Response.Write "</BODY></HTML>"
	'End the program so no further code will be executed
	Response.End
end if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This If statement will check to see if add is in the querystring
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request.QueryString("action") = "Add" then
if IsObject(Session("ShopCart")) then	'Check to see if we have any previously saved session variables
	set Basket = Session("ShopCart")	'Since we do, we set that info to our Basket
else									'Else if we dont have anythign else saved,
	Set Basket = CreateObject("Scripting.Dictionary")	'Create a new dictionary object
end if
dim Cnt
Cnt = Basket.Count
do
	'if Basket.Exists(Cnt) then
		Cnt = Cnt + 1
	'end if
loop until Basket.Exists("X" & Cnt) = false
Basket.Add "X" & Cnt, Request.Form("ITEM")		'Add and item to our basket, we take the total number
													'of items and add one
set Session("ShopCart") = Basket						'Save our session so we can use it later
tmpItems = Basket.Items									'Set tmpItems to hold all of our items in the basket
for i = 0 to Basket.Count - 1							'Loop to show whats currently in our basket
	'writes the value of the current item(i) and gives the option to delete it
	Response.Write i + 1 & ": " & tmpItems(i) & " - <a href='./Basket.asp?action=Del&Item=" & i & "'>Delete</a><BR>"
next
end if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This If statement will check to see if del is in the querystring
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if Request.QueryString("action") = "Del" then
if IsObject(Session("ShopCart")) then	'Check to see if we have any previously saved session variables
	set Basket = Session("ShopCart")	'Since we do, we set that info to our Basket
else									'Else if we dont have anythign else saved,
	Set Basket = CreateObject("Scripting.Dictionary")	'Create a new dictionary object
end if
tmpKeys = Basket.Keys
on error resume next									'Use this for error checking
Basket.Remove tmpKeys(int(trim(Request.QueryString("Item"))))						'Remove the item
if err.number <> 0 then									'If error other than 0
	Response.Write "Error " & err.number & "<BR>" & err.Description & "<P>"			'display error info
	Response.Write "QueryString Item: " & trim(Request.QueryString("Item")) & "<P>" 'Display
end if
tmpItems = Basket.Items									'Set tmpItems to hold all of our items in the basket
set Session("ShopCart") = Basket						'Save our session so we can use it later
tmpItems = Basket.Items									'Set tmpItems to hold all of our items in the basket
for i = 0 to Basket.Count - 1							'Loop to show whats currently in our basket
	'writes the value of the current item(i) and gives the option to delete it
	Response.Write i + 1 & ": " & tmpItems(i) & " - <a href='./Basket.asp?action=Del&Item=" & i & "'>Delete</a><BR>"
next
end if
%>
```

