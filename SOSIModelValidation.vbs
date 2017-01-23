option explicit 
 
 !INC Local Scripts.EAConstants-VBScript 
 
' 
' This script contains code from the default Project Browser template. 
' If you wish to modify this template, it is located in the Config\Script Templates 
' directory of your EA install path.    
' 
' Script Name: SOSI model validation 
' Author: Section for technology and standardization - Norwegian Mapping Authority
' Version: 1.1
' Date: 2017-01-23 
' Purpose: Validate model elements according to rules defined in the standard SOSI Regler for UML-modellering 5.0 
' Implemented rules: 
'	/krav/3:  
'			Find elements (classes, attributes, navigable association roles, operations, datatypes)  
'	        without definition (notes/rolenotes) in the selected package and subpackages 
'
'	/krav/6:		
'			Iso 19103 Requirement 6 - NCNames in codelist codes.
' 	/krav/7:	    
'			Iso 19103 Requirement 7 - definition of codelist codes.
'  	/krav/10: 
'			Check if all navigable association ends have cardinality 
'	/krav/11: 
'			Check if all navigable association ends have role names 
'	/krav/12: 
'			If datatypes have associations then the datatype shall only be target in a composition 
'  	/krav/14:
'			Iso 19103 Requirement 14 -inherit from same stereotypes
'  	/krav/15:
'			Iso 19103 Requirement 15 -known stereotypes
'  	/krav/16
'			Iso 19103 Requirement 16 -legal NCNames case-insesnitively unique within their namespace
'  	/krav/18
'			Iso 19103 Requirement 18 -all elements shall show all structures in at least one diagram
'			Current version test all classes and their attributes in diagrams, not yet roles and inheritance.
'	/krav/definisjoner: 
'			Same as krav/3 but checks also for definitions of packages and constraints
'			The part that checks definitions of constraints is implemented in sub checkConstraint	
'			The rest is implemented in sub checkDefinitions
'	/krav/eksternKodeliste
' 			Check if the coedlist has an asDictionary with value "true", if so, checks if the taggedValue "codeList" exist and if the value is valid or not.
'			Some parts missing. 2 subs.
'	/krav/enkelArv
' 			To check each class for multiple inheritance 
'	/krav/flerspråklighet/element:		
' 			if tagged value: "designation", "description" or "definition" exists, the value of the tag must end with "@<language-code>". 
' 			Checks attributes, operations, (roles), (constraints) and objecttypes 
'	/krav/flerspråklighet/pakke:
'			Check if the ApplicationSchema-package got a tagged value named "language" (error message if that is not the case) 
'			and if the value of it is empty or not (error message if empty). 
' 			And if there are designation-tags, checks that they have correct structure: "{name}"@{language}
' 	/krav/hoveddiagram/detaljering/navnining 
'			Check if a package with stereotype applicationSchema has more than one diagram called "Hoveddiagram", if so, checks that theres more characters
' 			in the name after "Hoveddiagram". If there is several "Hoveddiagram"s and one or more diagrams just named "Hoveddiagram" it returns an error. 
'  	/krav/hoveddiagram/navning: 
'			Check if an application-schema has less than one diagram named "Hoveddiagram", if so, returns an error 
' 	/krav/oversiktsdiagram:
'			Check that a package with more than one diagram with name starting with "Hoveddiagram" also has at least one diagram called "Oversiktsdiagram" 
'	/krav/navning (partially): 
'			Check if names of attributes, operations, roles start with lower case and names of packages,  
'			classes and associations start with upper case 
'	/krav/SOSI-modellregister/applikasjonsskjema/status
'			Check if the ApplicationSchema-package got a tagged value named "SOSI_modellstatus" and checks if it is a valid value
'   /krav/SOSI-modellregister/applikasjonsskjema/versjonsnummer
'           Check if the last part of the package name is a version number.  Ignores the text "Utkast" for this check
'   /krav/SOSI-modellregister/applikasjonsskjema/standard/pakkenavn/utkast
'			Check if packages with SOSI_modellstatus tag "utkast" has "Utkast" in package name. Also do the reverse check.
'  	/req/uml/constraint
'			To check if a constraint lacks name or definition. 
'  	/req/uml/packaging:
'     		To check if the value of the version-tag (tagged values) for an ApplicationSchema-package is empty or not. 
'   /anbefaling/1:
'			Checks every initial values in codeLists and enumerations for a package. If one or more initial values are numeric in one list, 
' 			it return a warning message. 
'  	/anbefaling/styleGuide:
'			Checks that the stereotype for packages and elements got the right use of lower- and uppercase, if not, return an error. Stereotypes to be cheked:
'			CodeList, dataType, enumeration, interface, Leaf, Union, FeatureType, ApplicationSchema
'	/req/uml/profile      
'			from iso 19109 -well known types for all attributes, including iso 19103 Requirement 22 and 25
'	/req/uml/feature
'			featureType classes shall have unique names within the applicationSchema		
'	/krav/taggedValueSpråk 	
'			ApplicationSchema packages shall have a language tag, designation tag and definition tag. Partially implemented, does not check definition tag
'
' 
'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Project Browser Script main function 
 
 sub OnProjectBrowserScript() 
 	 
	Repository.EnsureOutputVisible("Script")
 	' Get the type of element selected in the Project Browser 
 	dim treeSelectedType 
 	treeSelectedType = Repository.GetTreeSelectedItemType() 
 	 
 	' Handling Code: Uncomment any types you wish this script to support 
 	' NOTE: You can toggle comments on multiple lines that are currently 
 	' selected with [CTRL]+[SHIFT]+[C]. 
 	select case treeSelectedType 
 	 
 '		case otElement 
 '			' Code for when an element is selected 
 '			dim theElement as EA.Element 
 '			set theElement = Repository.GetTreeSelectedObject() 
 '					 
 		case otPackage 
 			' Code for when a package is selected 
 			dim thePackage as EA.Package 
 			set thePackage = Repository.GetTreeSelectedObject() 
 			'check if the selected package has stereotype applicationSchema 
 			if UCase(thePackage.element.stereotype) = UCase("applicationSchema") then 
				
				dim box, mess
				'mess = 	"Model validation 2016-08-19 Logging errors and warnings."&Chr(13)&Chr(10)
				mess = "Model validation based on requirements and recommendations in SOSI standard 'Regler for UML-modellering 5.0'"&Chr(13)&Chr(10)
				mess = mess + ""&Chr(13)&Chr(10)
				mess = mess + "Please find a list with the implemented rules in this script's source code (line 15++)."&Chr(13)&Chr(10)
				'mess= mess +  "/krav/3 - elements with definition."&Chr(13)&Chr(10)
				'mess = mess + "/krav/definisjoner - packages and constraints with definition."&Chr(13)&Chr(10)
				'mess = mess + "/krav/6 (Iso 19103 Req 6) - NCNames for codes."&Chr(13)&Chr(10)
				'mess = mess + "/krav/7 (Iso 19103 Req 7) - definition on codes."&Chr(13)&Chr(10)
				'mess = mess + "/krav/10	(Iso 19103 Req 10) - multiplicity."&Chr(13)&Chr(10)
				'mess = mess + "/krav/11	(Iso 19103 Req 11) - role names."&Chr(13)&Chr(10)
				'mess = mess + "/krav/flerspråklighet/pakke - tagged value 'language'."&Chr(13)&Chr(10)
				'mess = mess + "/krav/12	(Iso 19103 Req 12) - datatypes target in composition."&Chr(13)&Chr(10)
				'mess = mess + "/krav/enkelArv - single inheritance."&Chr(13)&Chr(10)
				'mess = mess + "/krav/Navning - all names CamelCase."&Chr(13)&Chr(10)
				'mess = mess + "/anbefaling/1 (Iso 19103 Rec 1) - meaningful initial values."&Chr(13)&Chr(10)
				'mess = mess + "/req/uml/packaging (Iso 19109) - tagged value 'version'."&Chr(13)&Chr(10)
				'mess = mess + "/krav/SOSI-modellregister - known SOSI model registry status codes."&Chr(13)&Chr(10)
				'mess = mess + "/krav/14	(Iso 19103 Req 14) - inherit from same stereotypes."&Chr(13)&Chr(10)
				'mess = mess + "/krav/15	(Iso 19103 Req 15) - known stereotypes."&Chr(13)&Chr(10)
				'mess = mess + "/krav/16	(Iso 19103 Req 16) - legal NCNames case-insensitively unique."&Chr(13)&Chr(10)
				'mess = mess + "/req/uml/profile	(Iso 19103, 19107, 19109) - well known types."&Chr(13)&Chr(10)
				mess = mess + ""&Chr(13)&Chr(10)
				mess = mess + "Starts model validation for package [" & thePackage.Name &"]."&Chr(13)&Chr(10)

				box = Msgbox (mess, vbOKCancel, "SOSI model validation 1.0.8")
				select case box
					case vbOK
						'inputBoxGUI to receive user input regarding the log level
						dim logLevelFromInputBox, logLevelInputBoxText, correctInput, abort
						logLevelInputBoxText = "Please select the log level."&Chr(13)&Chr(10)
						logLevelInputBoxText = logLevelInputBoxText+ ""&Chr(13)&Chr(10)
						logLevelInputBoxText = logLevelInputBoxText+ ""&Chr(13)&Chr(10)
						logLevelInputBoxText = logLevelInputBoxText+ "E - Error log level: logs error messages only."&Chr(13)&Chr(10)
						logLevelInputBoxText = logLevelInputBoxText+ ""&Chr(13)&Chr(10)
						logLevelInputBoxText = logLevelInputBoxText+ "W - Warning log level (recommended): logs error and warning messages."&Chr(13)&Chr(10)
						logLevelInputBoxText = logLevelInputBoxText+ ""&Chr(13)&Chr(10)
						logLevelInputBoxText = logLevelInputBoxText+ "Enter E or W:"&Chr(13)&Chr(10)
						correctInput = false
						abort = false
						do while not correctInput
						
							logLevelFromInputBox = InputBox(logLevelInputBoxText, "Select log level", "W")
							select case true 
								case UCase(logLevelFromInputBox) = "E"	
									'code for when E = Error log level has been selected, only Error messages will be shown in the Script Output window
									globalLogLevelIsWarning = false
									correctInput = true
								case UCase(logLevelFromInputBox) = "W"	
									'code for when W = Error log level has been selected, both Error and Warning messages will be shown in the Script Output window
									globalLogLevelIsWarning = true
									correctInput = true
								case IsEmpty(logLevelFromInputBox)
									'user pressed cancel or closed the dialog
									MsgBox "Abort",64
									abort = true
									exit do
								case else
									MsgBox "You made an incorrect selection! Please enter either 'E' or 'W'.",48
							end select
						
						loop
						
						if not abort then
							'For /krav/18:
							set startPackage = thePackage
							Set diaoList = CreateObject( "System.Collections.Sortedlist" )
							Set diagList = CreateObject( "System.Collections.Sortedlist" )
							recListDiagramObjects(thePackage)

							Dim StartTime, EndTime, Elapsed
							StartTime = timer 
							startPackageName = thePackage.Name
							FindInvalidElementsInPackage(thePackage) 
							Elapsed = formatnumber((Timer - StartTime),2)
							'------------------------------------------------------------------ 
							'---Check global variables--- 
							'------------------------------------------------------------------ 
	
							'check uniqueness of featureType names
							checkUniqueFeatureTypeNames()
	
							'error-message for /krav/hoveddiagram/navning (sub procedure: CheckPackageForHoveddiagram)
							'if the applicationSchema package got less than one diagram with a name starting with "Hoveddiagram", then return an error 	
							if 	not foundHoveddiagram  then
								Session.Output("Error: Neither package [" &startPackageName& "] nor any of it's subpackages has a diagram with a name starting with 'Hoveddiagram' [/krav/hoveddiagram/navning]")
								globalErrorCounter = globalErrorCounter + 1 
					
							end if 	
							
							'error-message for /krav/hoveddiagram/detaljering/navning (sub: FindHoveddiagramsInAS)
							'if the applicationSchema package got more than one diagram named "Hoveddiagram", then return an error 
							if numberOfHoveddiagram > 1 or (numberOfHoveddiagram = 1 and numberOfHoveddiagramWithAdditionalInformationInTheName > 0) then 
								dim sumOfHoveddiagram 
								sumOfHoveddiagram = numberOfHoveddiagram + numberOfHoveddiagramWithAdditionalInformationInTheName
								Session.Output("Error: Package ["&startPackageName&"] has "&sumOfHoveddiagram&" diagrams named 'Hoveddiagram' and "&numberOfHoveddiagram&" of them named exactly 'Hoveddiagram'. When there are multiple diagrams of that type additional information is expected in the diagrams' name. [/krav/hoveddiagram/detaljering/navning]")
								globalErrorCounter = globalErrorCounter + 1 
			
							end if 
	
							
							Session.Output("Number of errors found: " & globalErrorCounter) 
							if globalLogLevelIsWarning then
								Session.Output("Number of warnings found: " & globalWarningCounter)
							end if	
							Session.Output("Run time: " &Elapsed& " seconds" )
						end if	
					case VBcancel
						'nothing to do						
				end select 
			else 
 				Msgbox "Package [" & thePackage.Name &"] does not have stereotype «ApplicationSchema». Select a package with stereotype «ApplicationSchema» to start model validation." 
 			end if 
 			 
 			 
 			 
'		case otDiagram 
'			' Code for when a diagram is selected 
'			dim theDiagram as EA.Diagram 
'			set theDiagram = Repository.GetTreeSelectedObject() 
'			 
'		case otAttribute 
'			' Code for when an attribute is selected 
'			dim theAttribute as EA.Attribute 
'			set theAttribute = Repository.GetTreeSelectedObject() 
'			 
'		case otMethod 
'			' Code for when a method is selected 
'			dim theMethod as EA.Method 
'			set theMethod = Repository.GetTreeSelectedObject() 
 		 
 		case else 
 			' Error message 
 			Session.Prompt "[Warning] You must select a package with stereotype ApplicationSchema in the Project Browser to start the validation.", promptOK 
 			 
 	end select 
 	 
end sub 
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------
 
 
'------------------------------------------------------------START-------------------------------------------------------------------------------------------
'Sub name: 		CheckDefinition
'Author: 		Magnus Karge
'Date: 			20160925 
'Purpose: 		Check if the provided argument for input parameter theObject fulfills the requirements in [krav/3]: 
'				Find elements (classes, attributes, navigable association roles, operations, datatypes)  
'				without definition (notes/rolenotes) 
'				and [krav/definisjoner]: 
'				Find packages and constraints without definition
'@param[in] 	theObject (EA.ObjectType) The object to check,  
'				supposed to be one of the following types: EA.Attribute, EA.Method, EA.Connector, EA.Element 
 
 sub CheckDefinition(theObject) 
 	'Declare local variables 
 	Dim currentAttribute as EA.Attribute 
 	Dim currentMethod as EA.Method 
 	Dim currentConnector as EA.Connector 
 	Dim currentElement as EA.Element 
	Dim currentPackage as EA.Package
 		 
 	Select Case theObject.ObjectType 
 		Case otElement 
 			' Code for when the function's parameter is an element 
 			set currentElement = theObject 
 			 
 			If currentElement.Notes = "" then 
 				Session.Output("Error: Class [«" &getStereotypeOfClass(currentElement)& "» "& currentElement.Name & "] has no definition. [/krav/3] & [/krav/definisjoner]")	 
 				globalErrorCounter = globalErrorCounter + 1 
 			end if 
 		Case otAttribute 
 			' Code for when the function's parameter is an attribute 
 			 
 			set currentAttribute = theObject 
 			 
 			'get the attribute's parent element 
 			dim attributeParentElement as EA.Element 
 			set attributeParentElement = Repository.GetElementByID(currentAttribute.ParentID) 
 			 
 			if currentAttribute.Notes = "" then 
				Session.Output( "Error: Class [«" &getStereotypeOfClass(attributeParentElement)& "» "& attributeParentElement.Name &"] \ attribute [" & currentAttribute.Name & "] has no definition. [/krav/3] & [/krav/definisjoner]") 
 				globalErrorCounter = globalErrorCounter + 1 
 			end if 
 			 
 		Case otMethod 
 			' Code for when the function's parameter is a method 
 			 
 			set currentMethod = theObject 
 			 
 			'get the method's parent element, which is the class the method is part of 
 			dim methodParentElement as EA.Element 
 			set methodParentElement = Repository.GetElementByID(currentMethod.ParentID) 
 			 
 			if currentMethod.Notes = "" then 
 				Session.Output( "Error: Class [«" &getStereotypeOfClass(methodParentElement)& "» "& methodParentElement.Name &"] \ operation [" & currentMethod.Name & "] has no definition. [/krav/3] & [/krav/definisjoner]") 
 				globalErrorCounter = globalErrorCounter + 1 
 			end if 
 		Case otConnector 
 			' Code for when the function's parameter is a connector 
 			 
 			set currentConnector = theObject 
 			 
 			'get the necessary connector attributes 
 			dim sourceEndElementID 
 			sourceEndElementID = currentConnector.ClientID 'id of the element on the source end of the connector 
 			dim sourceEndNavigable  
 			sourceEndNavigable = currentConnector.ClientEnd.Navigable 'navigability on the source end of the connector 
 			dim sourceEndName 
 			sourceEndName = currentConnector.ClientEnd.Role 'role name on the source end of the connector 
 			dim sourceEndDefinition 
 			sourceEndDefinition = currentConnector.ClientEnd.RoleNote 'role definition on the source end of the connector 
 								 
 			dim targetEndNavigable  
 			targetEndNavigable = currentConnector.SupplierEnd.Navigable 'navigability on the target end of the connector 
 			dim targetEndName 
 			targetEndName = currentConnector.SupplierEnd.Role 'role name on the target end of the connector 
 			dim targetEndDefinition 
 			targetEndDefinition = currentConnector.SupplierEnd.RoleNote 'role definition on the target end of the connector 
 
 
 			dim sourceEndElement as EA.Element 
 			 
 			if sourceEndNavigable = "Navigable" and sourceEndDefinition = "" then 
 				'get the element on the source end of the connector 
 				set sourceEndElement = Repository.GetElementByID(sourceEndElementID) 
 				 
				Session.Output( "Error: Class [«" &getStereotypeOfClass(sourceEndElement)& "» "& sourceEndElement.Name &"] \ association role [" & sourceEndName & "] has no definition. [/krav/3] & [/krav/definisjoner]") 
 				globalErrorCounter = globalErrorCounter + 1 
 			end if 
 			 
 			if targetEndNavigable = "Navigable" and targetEndDefinition = "" then 
 				'get the element on the source end of the connector (also source end element here because error message is related to the element on the source end of the connector) 
 				set sourceEndElement = Repository.GetElementByID(sourceEndElementID) 
 				 
				Session.Output( "Error: Class [«"&getStereotypeOfClass(sourceEndElement)&"» "&sourceEndElement.Name &"] \ association role [" & targetEndName & "] has no definition. [/krav/3] & [/krav/definisjoner]") 
 				globalErrorCounter = globalErrorCounter + 1 
 			end if 
 		Case otPackage 
 			' Code for when the function's parameter is a package 
 			 
 			set currentPackage = theObject 
 			 
 			'check package definition 
			if currentPackage.Notes = "" then 
				Session.Output("Error: Package [" & currentPackage.Name & "] lacks a definition. [/krav/definisjoner]") 
				globalErrorCounter = globalErrorCounter + 1 
			end if 	 
 		Case else		 
 			'TODO: need some type of exception handling here
			Session.Output( "Debug: Function [CheckDefinition] started with invalid parameter.") 
 	End Select 
 	 
end sub 
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
'Purpose: 		help function in order to set stereotype that is shown 
'				in diagrams but not accessible as such via EAObjectAPI
'Used in sub: 	checkElementName
'@param[in]: theClass (EA.Element)
'returns: theClass's visible stereotype as character string, empty string if nothing found
 function getStereotypeOfClass(theClass)
	dim visibleStereotype
	visibleStereotype = ""
	if (Ucase(theClass.Stereotype) = Ucase("featuretype")) OR (Ucase(theClass.Stereotype) = Ucase("codelist")) OR (Ucase(theClass.Stereotype) = Ucase("datatype")) OR (Ucase(theClass.Stereotype) = Ucase("enumeration")) then
		'param theClass is Classifier subtype Class with different stereotypes
		visibleStereotype = theClass.Stereotype
	elseif (Ucase(theClass.Type) = Ucase("enumeration")) OR (Ucase(theClass.Type) = Ucase("datatype"))  then
		'param theClass is Classifier subtype DataType or Enumeration
		visibleStereotype = theClass.Type
	end if
	getStereotypeOfClass=visibleStereotype
 end function
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------
 
 
'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub name: checkElementName
' Author: Magnus Karge
' Date: 20160925 
' Purpose:  sub procedure to check if a given element's name is written correctly
' 			Implementation of /krav/navning
' 			
' @param[in]: theElement (EA.Element). The element to check. Can be class, enumeration, data type, attribute, operation, association, role or package
 
sub checkElementName(theElement) 
	
	select case theElement.ObjectType
		case otPackage
			'sub parameter is ObjectType oTPackage, check if first letter of the package's name is a capital letter 
 			if not Left(theElement.Name,1) = UCase(Left(theElement.Name,1)) then 
				Session.Output("Error: Package name [" & theElement.Name & "] shall start with capital letter. [/krav/navning]") 
				globalErrorCounter = globalErrorCounter + 1 
 			end if
		case otElement
			'sub's parameter is ObjectType oTElement, check if first letter of the element's name is a capital letter (element covers class, enumeration, datatype)
 			if not Left(theElement.Name,1) = UCase(Left(theElement.Name,1)) then 
 				Session.Output("Error: Class name [«"&getStereotypeOfClass(theElement)&"» "& theElement.Name & "] shall start with capital letter. [/krav/navning]") 
 				globalErrorCounter = globalErrorCounter + 1 
 			end if 
		case otAttribute
			'sub's parameter is ObjectType oTAttribute, check if first letter of the attribute's name is NOT a capital letter 
			if not Left(theElement.Name,1) = LCase(Left(theElement.Name,1)) then 
				dim attributeParentElement as EA.Element
				set attributeParentElement = Repository.GetElementByID(theElement.ParentID)
				Session.Output("Error: Attribute name [" & theElement.Name & "] in class [«"&getStereotypeOfClass(attributeParentElement)&"» "& attributeParentElement.Name &"] shall start with lowercase letter. [/krav/navning]") 
				globalErrorCounter = globalErrorCounter + 1
			end if									
 		case otConnector
			dim connector as EA.Connector
			set connector = theElement
			'sub's parameter is ObjectType oTConnector, check if the association has a name (not necessarily the case), if so check if the name starts with a capital letter 
			if not (connector.Name = "" OR len(connector.Name)=0) and not Left(connector.Name,1) = UCase(Left(connector.Name,1)) then 
				dim associationSourceElement as EA.Element
				dim associationTargetElement as EA.Element
				set associationSourceElement = Repository.GetElementByID(connector.ClientID)
				set associationTargetElement = Repository.GetElementByID(connector.SupplierID)
				Session.Output("Error: Association name [" & connector.Name & "] between class [«"&getStereotypeOfClass(associationSourceElement)&"» "& associationSourceElement.Name &"] and class [«"&getStereotypeOfClass(associationTargetElement)&"» " & associationTargetElement.Name & "] shall start with capital letter. [/krav/navning]") 
				globalErrorCounter = globalErrorCounter + 1 
			end if 
		'case otOperation
		'case otRole
	end select	
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub name: findMultipleInheritance
' Author: Sara Henriksen
' Date: 14.07.16 
' Purpose:  sub procedure to check if a given class has multiple inheritance 
' 			Implementation of /krav/enkelArv
' 			
' @param[in]: currentElement (EA.Element). The "class" to check 
 
sub findMultipleInheritance(currentElement) 
 
	loopCounterMultipleInheritance = loopCounterMultipleInheritance + 1 
 	dim connectors as EA.Collection  
  	set connectors = currentElement.Connectors  
  					  
  	'iterate the connectors  
  					 
  	dim connectorsCounter  
 	dim numberOfSuperClasses  
 	numberOfSuperClasses = 0  
 	dim theTargetGeneralization as EA.Connector 
 	set theTargetGeneralization = nothing 
 					 
	for connectorsCounter = 0 to connectors.Count - 1  
		dim currentConnector as EA.Connector  
		set currentConnector = connectors.GetAt( connectorsCounter )  
 						 
 						 
		'check if the connector type is "Generalization" and if so 
		'get the element on the source end of the connector   
		if currentConnector.Type = "Generalization"  then 
			if currentConnector.ClientID = currentElement.ElementID then  
 					 
				'count number of classes with a generalization connector on the source side  
				numberOfSuperClasses = numberOfSuperClasses + 1  
				set theTargetGeneralization = currentConnector  
			end if  
		end if 

		'if theres more than one generalization connecter on the source side the class has multiple inheritance 
		if numberOfSuperClasses > 1 then 
			Session.Output("Error: Class [«"&startClass.Stereotype&"» "&startClass.Name& "] has multiple inheritance. [/krav/enkelarv]") 
			globalErrorCounter = globalErrorCounter + 1 
			exit for  
		end if  
	next 
 					 
	' if there is just one generalization connector on the source side, start checking genralization connectors for the superclasses  
	' stop if number of loops exceeds 20
	if numberOfSuperClasses = 1 and not theTargetGeneralization is nothing and loopCounterMultipleInheritance < 21 then 
 				
		dim superClassID  
		dim superClass as EA.Element 
		'the elementID of the element at the target end 
		superClassID =  theTargetGeneralization.SupplierID  
		set superClass = Repository.GetElementByID(superClassID) 

		'Check level of superClass 
		call findMultipleInheritance (superClass) 
		elseif loopCounterMultipleInheritance = 21 then 
			Session.Output("Warning: Found more than 20 inheritance levels for class: [" &startClass.Name& "] while testing [/krav/enkelarv]. Please check for possible circle inheritance")
			globalWarningCounter = globalWarningCounter + 1 
	end if  
 end sub 
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Script Name: checkTVLanguageAndDesignation
' Author: Sara Henriksen
' Date: 26.07.16
' Purpose: Check if the ApplicationSchema-package got a tag named "language" and  check if the value is empty or not. 
' And if there is a designation tag, checks that it has correct structure: "{name}"@{language}  
' /krav/flersprålighet/pakke	
' sub procedure to check if the package has the provided tags with a value with correct structure
' @param[in]: theElement (Package Class) and taggedValueName (String)

sub checkTVLanguageAndDesignation(theElement, taggedValueName)

	if taggedValueName = "language" then 
 		if UCase(theElement.Stereotype) = UCase("applicationSchema") then
		
			dim packageTaggedValues as EA.Collection 
 			set packageTaggedValues = theElement.TaggedValues 

 			dim taggedValueLanguageMissing 
 			taggedValueLanguageMissing = true 
			'iterate trough the tagged values 
 			dim packageTaggedValuesCounter 
 			for packageTaggedValuesCounter = 0 to packageTaggedValues.Count - 1 
 				dim currentTaggedValue as EA.TaggedValue 
 				set currentTaggedValue = packageTaggedValues.GetAt(packageTaggedValuesCounter) 
				
				'check if the provided tagged value exist
				if (currentTaggedValue.Name = "language") and not (currentTaggedValue.Value= "") then 
					'check if the value is no or en, if not, retrun a warning 
					if not mid(StrReverse(currentTaggedValue.Value),1,2) = "ne" and not mid(StrReverse(currentTaggedValue.Value),1,2) = "on" then	
						if globalLogLevelIsWarning then
							Session.Output("Warning: Package [«"&theElement.Stereotype&"» " &theElement.Name&"] \ tag ["&currentTaggedvalue.Name& "] has a value which is not <no> or <en>. [/krav/flerspråklighet/pakke][/krav/taggedValueSpråk]")
							globalWarningCounter = globalWarningCounter + 1 
						end if
					end if
					taggedValueLanguageMissing = false 
					exit for 
				end if   
				if currentTaggedValue.Name = "language" and currentTaggedValue.Value= "" then 
					Session.Output("Error: Package [«"&theElement.Stereotype&"» " &theElement.Name&"] \ tag ["& currentTaggedValue.Name &"] lacks a value. [/krav/flerspråklighet/pakke][/krav/taggedValueSpråk]") 
					globalErrorCounter = globalErrorCounter + 1 
					taggedValueLanguageMissing = false 
					exit for 
				end if 
 			next 
			if taggedValueLanguageMissing then 
				Session.Output("Error: Package [«"&theElement.Stereotype&"» " &theElement.Name&"] lacks a [language] tag. [/krav/flerspråklighet/pakke][/krav/taggedValueSpråk]") 
				globalErrorCounter = globalErrorCounter + 1 
			end if 
		end if 
	end if 

	if taggedValueName = "designation" then

		if not theElement is nothing and Len(taggedValueName) > 0 then
		
			'check if the element has a tagged value with the provided name
			dim currentExistingTaggedValue1 AS EA.TaggedValue 
			dim valueExists
			dim enDesignation
			dim checkQuoteMark
			dim checkAtMark
			dim taggedValuesCounter1
			valueExists=false
			enDesignation = false
			for taggedValuesCounter1 = 0 to theElement.TaggedValues.Count - 1
				set currentExistingTaggedValue1 = theElement.TaggedValues.GetAt(taggedValuesCounter1)

				'check if the tagged value exists, and checks if the value starts with " and ends with "@{language}, if not, return an error. 
				if currentExistingTaggedValue1.Name = taggedValueName then
					valueExists=true
					checkQuoteMark=false
					checkAtMark=false
					
					if not len(currentExistingTaggedValue1.Value) = 0 then 

						if (InStr(currentExistingTaggedValue1.Value, "@en")<>0) then 
							enDesignation=true
						end if
						
						if (mid(currentExistingTaggedValue1.Value, 1, 1) = """") then 
							checkQuoteMark=true
						end if
						if (InStr(currentExistingTaggedValue1.value, """@")<>0) then 
							checkAtMark=true
						end if
						
						if not (checkAtMark and checkQuoteMark) then
							globalErrorCounter = globalErrorCounter + 1 
						end if 
					
						'Check if the value contains  illegal quotation marks, gives an Warning-message  
						dim startContent, endContent, designationContent
	
						startContent = InStr( currentExistingTaggedValue1.Value, """" ) 			
						endContent = len(currentExistingTaggedValue1.Value)- InStr( StrReverse(currentExistingTaggedValue1.Value), """" ) -1
						if endContent<0 then endContent=0
						designationContent = Mid(currentExistingTaggedValue1.Value,startContent+1,endContent)				

						if InStr(designationContent, """") then
							if globalLogLevelIsWarning then
								Session.Output("Warning: Package [«" &theElement.Stereotype& "» " &theElement.Name&"] \ tag [designation] has a value ["&currentExistingTaggedValue1.Value&"] that contains illegal use of quotation marks.")
								globalWarningCounter = globalWarningCounter + 1 
							end if	
						end if
					else
						Session.Output("Error: Package [«" &theElement.Stereotype& "» " &theElement.Name& "] \ tag [designation] has no value [/krav/taggedValueSpråk]") 
						globalErrorCounter = globalErrorCounter + 1
					end if
				end if 						
			next
			if UCase(theElement.Stereotype) = UCase("applicationSchema") then
				if not valueExists then
					Session.Output("Error: Package [«"&theElement.Stereotype&"» " &theElement.Name&"] does not have a designation tag [/krav/taggedValueSpråk]")
					globalErrorCounter = globalErrorCounter + 1
				else
					if not enDesignation then
						Session.Output("Error: Package [«"&theElement.Stereotype&"» " &theElement.Name&"] \ tag [designation] lacks a value for English. Expected value ""{English designation}""@en [/krav/taggedValueSpråk]")
						globalErrorCounter = globalErrorCounter + 1
					end if
				end if
			end if
		end if 
	end if
end sub 
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Script Name: structurOfTVforElement
' Author: Sara Henriksen
' Date: 26.07.16	
' Purpose: Check that the value of a designation/description/definition tag got the structure “{value}”@{landcode}. 
' Implemented for objecttypes, attributes, roles and operations.
' Two subs, where structurOfTVforElement calls structureOfTVConnectorEnd if the parameter is a connector
' krav/flerspråklighet/element 
' sub procedure to find the provided tags for a connector, and if they exist, check the structure of the value.   
' @param[in]: theConnectorEnd (EA.Connector), taggedValueName (string) theConnectorEnd is potencially having tags: description, designation, definition, 
' with a value with wrong structure. 
sub structureOfTVConnectorEnd(theConnectorEnd,  taggedValueName)

	if not theConnectorEnd is nothing and Len(taggedValueName) > 0 then
	
		'check if the element has a tagged value with the provided name
		dim currentExistingTaggedValue as EA.RoleTag 
		dim taggedValuesCounter

		for taggedValuesCounter = 0 to theConnectorEnd.TaggedValues.Count - 1
			set currentExistingTaggedValue = theConnectorEnd.TaggedValues.GetAt(taggedValuesCounter)

			'if the tagged values exist, check the structure of the value 
			if currentExistingTaggedValue.Tag = taggedValueName then
				'check if the structure of the tag is: "{value}"@{languagecode}
				if not (mid(StrReverse(currentExistingTaggedValue.Value), 1,4)) = "ne@"""  and not (mid(StrReverse(currentExistingTaggedValue.Value), 1,4)) = "on@""" or not (mid((currentExistingTaggedValue.Value),1,1)) = """" then
					Session.Output("Error: Role [" &theConnectorEnd.Role& "] \ tag [" &currentExistingTaggedValue.Tag& "] has a value [" &currentExistingTaggedValue.Value& "] with wrong structure. Expected structure: ""{Name}""@{language}. [/krav/flerspråklighet/element]")
					globalErrorCounter = globalErrorCounter + 1 
				end if 
			end if 
		next
	end if 
end sub 
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
'sub procedure to find the provided tags and if they exist, check the structure of the value.   
'@param[in]: theElement (EA.ObjectType), taggedValueName (string) The object to check against krav/flerspråklighet/pakke,  
'supposed to be one of the following types: EA.Element, EA.Attribute, EA.Method, EA.Connector 
sub structurOfTVforElement (theElement, taggedValueName)

	if not theElement is nothing and Len(taggedValueName) > 0 and not theElement.ObjectType = otConnectorEnd   then

		'check if the element has a tagged value with the provided name
		dim currentExistingTaggedValue AS EA.TaggedValue 
		dim taggedValuesCounter

		for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
			set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)

			if currentExistingTaggedValue.Name = taggedValueName then
				'check the structure of the tag: "{value}"@{languagecode}
				if not (mid(StrReverse(currentExistingTaggedValue.Value), 1,4)) = "ne@"""  and not (mid(StrReverse(currentExistingTaggedValue.Value), 1,4)) = "on@""" or not (mid((currentExistingTaggedValue.Value),1,1)) = """" then
					Dim currentElement as EA.Element
					Dim currentAttribute as EA.Attribute
					Dim currentOperation as EA.Method
					
					Select Case theElement.ObjectType 
						'case element
						Case otElement 
							set currentElement = theElement 
						
							Session.Output("Error: Class [«"&theElement.Stereotype&"» " &theElement.Name& "] \ tag [" &currentExistingTaggedValue.Name& "] has a value [" &currentExistingTaggedValue.Value& "] with wrong structure. Expected structure: ""{Name}""@{language}. [/krav/flerspråklighet/element]")
							globalErrorCounter = globalErrorCounter + 1 
						
						'case attribute
						Case otAttribute
							set currentAttribute = theElement
						
							'get the element (class, enumeration, data Type) the attribute belongs to
							dim parentElementOfAttribute as EA.Element
							set parentElementOfAttribute = Repository.GetElementByID(currentAttribute.ParentID)
						
							Session.Output("Error: Class [«"& parentElementOfAttribute.Stereotype &"» "& parentElementOfAttribute.Name &"\ attribute [" &theElement.Name& "] \ tag [" &currentExistingTaggedValue.Name& "] has a value [" &currentExistingTaggedValue.Value& "] with wrong structure. Expected structure: ""{Name}""@{language}. [/krav/flerspråklighet/element]")
							globalErrorCounter = globalErrorCounter + 1 
						
						'case operation
						Case otMethod
							set currentOperation = theElement
							
							'get the element (class, enumeration, data Type) the operation belongs to
							dim parentElementOfOperation as EA.Element
							set parentElementOfOperation = Repository.GetElementByID(currentOperation.ParentID)
						
							Session.Output("Error: Class [«"& parentElementOfOperation.Stereotype &"» "& parentElementOfOperation.Name &"\ operation [" &theElement.Name& "] \ tag [" &currentExistingTaggedValue.Name& "] has a value: " &currentExistingTaggedValue.Value& " with wrong structure. Expected structure: ""{Name}""@{language}. [/krav/flerspråklighet/element]")
							globalErrorCounter = globalErrorCounter + 1 

					end select 	
				end if 
			end if 
		next
	'if the element is a connector then call another sub routine 
	elseif theElement.ObjectType = otConnectorEnd then
		Call structureOfTVConnectorEnd(theElement, taggedValueName)
	end if 
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Script Name: checkValueOfTVVersion
' Author: Sara Henriksen
' Date: 25.07.16 
' Purpose: To check if the value of the version-tag (tagged values) for an ApplicationSchema-package is empty or not. 
' req/uml/packaging
' sub procedure to check if the tagged value with the provided name exist in the ApplicationSchema, and if the value is emty it returns an Error-message. 
' @param[in]: theElement (Element Class) and TaggedValueName (String) 
sub checkValueOfTVVersion(theElement, taggedValueName)

	if UCase(theElement.stereotype) = UCase("applicationSchema") then

		if not theElement is nothing and Len(taggedValueName) > 0 then

			'check if the element has a tagged value with the provided name
			dim taggedValueVersionMissing
			taggedValueVersionMissing = true
			dim currentExistingTaggedValue AS EA.TaggedValue 
			dim taggedValuesCounter
			for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
				set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)
			
				'check if the taggedvalue exists, and if so, checks if the value is empty or not. An empty value will give an error-message. 
				if currentExistingTaggedValue.Name = taggedValueName then
					'remove spaces before and after a string, if the value only contains blanks  the value is empty
					currentExistingTaggedValue.Value = Trim(currentExistingTaggedValue.Value)
					if len (currentExistingTaggedValue.Value) = 0 then 
						Session.Output("Error: Package [«"&theElement.Stereotype&"» " &theElement.Name&"] has an empty version-tag. [req/uml/packaging]")
						globalErrorCounter = globalErrorCounter + 1 
						taggedValueVersionMissing = false 
					else
						taggedValueVersionMissing = false 
						'Session.Output("[" &theElement.Name& "] has version tag:  " &currentExistingTaggedValue.Value)
					end if 
				end if
			next
			'if tagged value version lacks for the package, return an error 
			if taggedValueVersionMissing then
				Session.Output ("Error: Package [«"&theElement.Stereotype&"» " &theElement.Name&"] lacks a [version] tag. [req/uml/packaging]")
				globalErrorCounter = globalErrorCounter + 1 
			end if
		end if 
	end if
end sub 
'-------------------------------------------------------------END-------------------------------------------------------------------------------------------- 
 

'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Script Name: checkConstraint
' Author: Sara Henriksen
' Date: 26.08.16
' Purpose: to check if a constraint lacks name or definition. 
' req/uml/constraint & krav/definisjoner
' sub procedure to check the current element/attribute/connector/package for constraints without name or definition
' not sure if it is possible in EA that constraints without names can exist, checking it anyways
' @param[in]: currentConstraint (EA.Constraint) theElement (EA.ObjectType) The object to check against req/uml/constraint,  
' supposed to be one of the following types: EA.Element, EA.Attribute, EA.Connector, EA.package

sub checkConstraint(currentConstraint, theElement)
	
	dim currentConnector as EA.Connector
	dim currentElement as EA.Element
	dim currentAttribute as EA.Attribute
	dim currentPackage as EA.Package
	
	Select Case theElement.ObjectType

		'if the object is an element
		Case otElement 
		set currentElement = theElement 
		
		'if the current constraint lacks definition, then return an error
		if currentConstraint.Notes= "" then 
			Session.Output("Error: Class [«"&theElement.Stereotype&"» "&theElement.Name&"] \ constraint [" &currentConstraint.Name&"] lacks definition. [/req/uml/constraint] & [krav/definisjoner]")
			globalErrorCounter = globalErrorCounter + 1 
		end if 
		
		'if the current constraint lacks a name, then return an error 
		if currentConstraint.Name = "" then
			Session.Output("Error: Class [«" &theElement.Stereotype& "» "&currentElement.Name& "] has a constraint without a name. [/req/uml/constraint]")
			globalErrorCounter = globalErrorCounter + 1 
		end if 
		
		'if the object is an attribute 
		Case otAttribute
		set currentAttribute = theElement 
		
		'if the current constraint lacks definition, then return an error
		dim parentElementID
		parentElementID = currentAttribute.ParentID
		dim parentElementOfAttribute AS EA.Element
		set parentElementOfAttribute = Repository.GetElementByID(parentElementID)
		if currentConstraint.Notes= "" then 
			Session.Output("Error: Class ["&parentElementOfAttribute.Name&"] \ attribute ["&theElement.Name&"] \ constraint [" &currentConstraint.Name&"] lacks definition. [/req/uml/constraint] & [krav/definisjoner]")
			globalErrorCounter = globalErrorCounter + 1 
		end if 
		
		'if the current constraint lacks a name, then return an error 	
		if currentConstraint.Name = "" then
			Session.Output("Error: Attribute ["&theElement.Name& "] has a constraint without a name. [/req/uml/constraint]")
			globalErrorCounter = globalErrorCounter + 1 
		end if 
		
		Case otPackage
		set currentPackage = theElement
		
		'if the current constraint lacks definition, then return an error message
		if currentConstraint.Notes= "" then 
			Session.Output("Error: Package [«"&theElement.Element.Stereotype&"» "&theElement.Name&"] \ constraint [" &currentConstraint.Name&"] lacks definition. [/req/uml/constraint] & [krav/definisjoner]")
			globalErrorCounter = globalErrorCounter + 1 
		end if 
		
		'if the current constraint lacks a name, then return an error meessage		
		if currentConstraint.Name = "" then
			Session.Output("Error: Package [«" &theElement.Element.Stereotype&"» " &currentElement.Name& "] has a constraint without a name. [/req/uml/constraint]")
			globalErrorCounter = globalErrorCounter + 1 
		end if 
			
		Case otConnector
		set currentConnector = theElement
		
		'if the current constraint lacks definition, then return an error message
		if currentConstraint.Notes= "" then 
		
			dim sourceElementID
			sourceElementID = currentConnector.ClientID
			dim sourceElementOfConnector AS EA.Element
			set sourceElementOfConnector = Repository.GetElementByID(sourceElementID)
			
			dim targetElementID
			targetElementID = currentConnector.SupplierID
			dim targetElementOfConnector AS EA.Element
			set targetElementOfConnector = Repository.GetElementByID(targetElementID)
		
			Session.Output("Error: Constraint [" &currentConstraint.Name&"] owned by connector [ "&theElement.Name&"] between class ["&sourceElementOfConnector.Name&"] and class ["&targetElementOfConnector.Name&"] lacks definition. [/req/uml/constraint] & [krav/definisjoner]")
			globalErrorCounter = globalErrorCounter + 1 
		end if 
		
		'if the current constraint lacks a name, then return an error message		
		if currentConstraint.Name = "" then
			Session.Output("Error: Connector [" &theElement.Name& "] has a constraint without a name. [/req/uml/constraint]")
			globalErrorCounter = globalErrorCounter + 1 
		
		end if
	end select
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------
 
 
'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Script Name: ValidValueSOSI_modellstatus 
' Author: Sara Henriksen
' Date: 25.07.16
' Purpose: Check if the ApplicationSchema-package got a tagged value named "SOSI_modellstatus" and checks if it is a valid value 
' /krav/SOSI-modellregister/applikasjonsskjema/status
' sub procedure to check if the tagged value with the provided name exist, and checks if the value is valid or not 
' (valid values: utkast, gyldig, utkastOgSkjult, foreslått, erstattet, tilbaketrukket og ugyldig). 
'@param[in]: theElement (Package Class) and TaggedValueName (String) 

sub ValidValueSOSI_modellstatus(theElement, taggedValueName)
	
	if UCase(theElement.Stereotype) = UCase("applicationSchema") then

		if not theElement is nothing and Len(taggedValueName) > 0 then
		
			'check if the element has a tagged value with the provided name
			dim taggedValueSOSIModellstatusMissing 
			taggedValueSOSIModellstatusMissing = true 
			dim currentExistingTaggedValue AS EA.TaggedValue 
			dim taggedValuesCounter
			
			for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
				set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)
			
				if currentExistingTaggedValue.Name = taggedValueName then
					'check if the value of the tag is one of the approved values. 
					if currentExistingTaggedValue.Value = "utkast" or currentExistingTaggedValue.Value = "gyldig" or currentExistingTaggedValue.Value = "utkastOgSkjult" or currentExistingTaggedValue.Value = "foreslått" or currentExistingTaggedValue.Value = "erstattet" or currentExistingTaggedValue.Value = "tilbaketrukket" or currentExistingTaggedValue.Value = "ugyldig" then 

						taggedValueSOSIModellstatusMissing = false 
					else
						Session.Output("Error: Package [«"&theElement.Stereotype&"» "&theElement.Name& "] \ tag [SOSI_modellstatus] has a value [" &currentExistingTaggedValue.Value& "]. The value is not approved. [/krav/SOSI-modellregister/applikasjonsskjema/status]")
						globalErrorCounter = globalErrorCounter + 1 
						taggedValueSOSIModellstatusMissing = false 
					end if 
				end if
			next

			'if the tag doesen't exist, return an error-message 
			if taggedValueSOSIModellstatusMissing then
				Session.Output("Error: Package [«"&theElement.Stereotype&"» " &theElement.Name& "] lacks a [SOSI_modellstatus] tag. [krav/SOSI-modellregister/applikansjonsskjema/status]")
				globalErrorCounter = globalErrorCounter + 1 
			end if 
		end if
	end if 
end sub 
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Script Name: checkNumericinitialValues
' Author: Sara Henriksen
' Date: 27.07.16
' Purpose: checks every initial values in  codeLists and enumerations for a package. Returns a warning for each attribute with intitial value that is numeric 
' /anbefaling/1
'sub procedure to check if the initial values of the attributes in a CodeList/enumeration are numeric or not. 
'@param[in]: theElement (EA.element) The element containing  attributes with potentially numeric inital values 
sub checkNumericinitialValues(theElement)

	dim attr as EA.Attribute
	dim numberOfNumericDefault

	'navigate through all attributes in the codeLists/enumeration 
	for each attr in theElement.Attributes 
		'check if the initial values are numeric 
		if IsNumeric(attr.Default)   then
			if globalLogLevelIsWarning then	
				Session.Output("Warning: Class [«"&theElement.Stereotype&"» "&theElement.Name&"] \ attribute [" &attr.Name& "] has numeric initial value [" &attr.Default& "] that is probably meaningless. Recommended to use script <flyttInitialverdiPåKodelistekoderTilSOSITag>. [/anbefaling/1]")		
				globalWarningCounter = globalWarningCounter + 1 
			end if
		end if 
	next
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Script Name: checkStereotypes
' Author: Sara Henriksen
' Date: 29.08.16 
'Purpose: check that the stereotype for packages and elements got the right use of lower- and uppercase, if not, return an error. Stereotypes to be cheked:
' CodeList, dataType, enumeration, interface, Leaf, Union, FeatureType, ApplicationSchema (case sensitiv)
' /anbefaling/styleGuide 
'sub procedure to check if the stereotype for a given package or element
'@param[in]: theElement (EA.ObjectType) The object to check against /anbefaling/styleguide 
'supposed to be one of the following types: EA.Element, EA.Package  

sub checkStereotypes(theElement)
	
	Dim currentElement as EA.Element
	Dim currentPackage as EA.Package

	Select Case theElement.ObjectType

		Case otPackage 
		set currentPackage = theElement 
		
		if UCase(theElement.Element.Stereotype) = "APPLICATIONSCHEMA" then
			if  not theElement.Element.Stereotype = "ApplicationSchema"   then 
				if globalLogLevelIsWarning then
					Session.Output("Warning: Package [«"&theElement.Element.Stereotype&"» "&theElement.Name&"]  has a stereotype with wrong use of lower-and uppercase. Expected use of case: ApplicationSchema [/anbefaling/styleGuide]")
					globalWarningCounter = globalWarningCounter + 1 
				end if	
			end if 
		end if 
	
		if UCase(theElement.Element.Stereotype) = "LEAF" then
			if  not theElement.Element.Stereotype = "Leaf" then 'and not pack.Element.Stereotype = "Leaf" then
				if globalLogLevelIsWarning then
					Session.Output("Warning: Package [«"&theElement.Element.Stereotype&" »"&theElement.Name&"]  has a stereotype with wrong use of lower-and uppercase. Expected use of case: Leaf [/anbefaling/styleGuide]")
					globalWarningCounter = globalWarningCounter + 1 
				end if	
			end if 
		end if
		
		Case otElement
		set currentElement = theElement 
		if UCase(theElement.Stereotype) = "CODELIST" then 
			if  not theElement.Stereotype = "CodeList" then 
				if globalLogLevelIsWarning then
					Session.Output("Warning: Element [«"&theElement.Stereotype&"» "&theElement.Name&"] has a stereotype with wrong use of lower-and uppercase. Expected use of case: CodeList [/anbefaling/styleGuide]")
					globalWarningCounter = globalWarningCounter + 1 
				end if	
			end if 
		end if 
		
		if UCase(theElement.Stereotype) = "DATATYPE" then 
			if  not theElement.Stereotype = "dataType" then 
				if globalLogLevelIsWarning then
					Session.Output("Warning: Element [«"&theElement.Stereotype&"» "&theElement.Name&"] has a stereotype with wrong use of lower-and uppercase. Expected use of case: dataType [/anbefaling/styleGuide]")
					globalWarningCounter = globalWarningCounter + 1 
				end if	
			end if 
		end if 
		
		if UCase(theElement.Stereotype) = "FEATURETYPE" then 
			if  not theElement.Stereotype = "FeatureType" then 
				if globalLogLevelIsWarning then
					Session.Output("Warning: Element [«"&theElement.Stereotype&"» "&theElement.Name&"] has a stereotype with wrong use of lower-and uppercase. Expected use of case: FeatureType [/anbefaling/styleGuide]")
					globalWarningCounter = globalWarningCounter + 1 
				end if	
			end if 
		end if 
		
		if UCase(theElement.Stereotype) = "UNION" then 
			if  not theElement.Stereotype = "Union" then 
				if globalLogLevelIsWarning then
					Session.Output("Warning: Element [«"&theElement.Stereotype&"» "&theElement.Name&"] has a stereotype with wrong use of lower-and uppercase. Expected use of case: Union [/anbefaling/styleGuide]")
					globalWarningCounter = globalWarningCounter + 1 
				end if	
			end if 
		end if
		
		if UCase(theElement.Stereotype) = "ENUMERATION" then 
			if  not theElement.Stereotype = "enumeration" then 
				if globalLogLevelIsWarning then
					Session.Output("Warning: Element [«"&theElement.Stereotype&"» "&theElement.Name&"] has a stereotype with wrong use of lower-and uppercase. Expected use of case: enumeration [/anbefaling/styleGuide]")
					globalWarningCounter = globalWarningCounter + 1 
				end if	
			end if 
		end if
		
		if UCase(theElement.Stereotype) = "INTERFACE" then 
			if  not theElement.Stereotype = "interface" then 
				if globalLogLevelIsWarning then
					Session.Output("Warning: Element [«"&theElement.Stereotype&"» "&theElement.Name&"] has a stereotype with wrong use of lower-and uppercase. Expected use of case: interface [/anbefaling/styleGuide]")
					globalWarningCounter = globalWarningCounter + 1 
				end if	
			end if 
		end if
	end select 
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Script Name: CheckPackageForHoveddiagram
' Author: Sara Henriksen
' Date: 03.08.16
' Purpose: Check if an application-schema has less than one diagram named "Hoveddiagram", if so, returns an error
' /krav/hoveddiagram/navning
'sub procedure to check if the given package got one or more diagrams with a name starting with "Hoveddiagram", if not, returns an error 
'@param[in]: package (EA.package) The package containing diagrams potentially with one or more names without "Hoveddiagram".
sub CheckPackageForHoveddiagram(package)
	
	dim diagrams as EA.Collection
	set diagrams = package.Diagrams
	'check all digrams in the package 
	dim i
	for i = 0 to diagrams.Count - 1
		dim currentDiagram as EA.Diagram
		set currentDiagram = diagrams.GetAt( i )
		'set foundHoveddiagram true if any diagrams have been found with a name starting with "Hoveddiagram"
		if Mid((currentDiagram.Name),1,12) = "Hoveddiagram"  then 
			foundHoveddiagram = true 
		end if	
	next
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Script Name: FindHoveddiagramsInAS
' Author: Sara Henriksen
' Date: 03.08.16
' Purpose: to check if the applicationSchema-package has more than one diagram with a name starting with "Hoveddiagram", if so, returns an error if the 
' name of the Diagram is nothing more than "Hoveddiagram". Returns one error per ApplicationSchema, with the number of wrong-named diagrams for the package.
' /krav/hoveddiagram/detaljering/navning 
' sub procedure to check if the given package and its subpackages has more than one diagram with the provided name, if so, return and error if 
' the name of the Diagram is nothing more than "Hoveddiagram".
'@param[in]: package (EA.package) The package potentially containing diagrams with the provided name

sub FindHoveddiagramsInAS(package)
	
	dim diagrams as EA.Collection
	set diagrams = package.Diagrams

	'find all digrams in the package 
	dim i
	for i = 0 to diagrams.Count - 1
		dim currentDiagram as EA.Diagram
		set currentDiagram = diagrams.GetAt( i )
				
		'if the package got less than one diagram with a name starting with "Hoveddiagram", then return an error 
		if UCase(Mid((currentDiagram.Name),1,12)) = "HOVEDDIAGRAM" and len(currentDiagram.Name) = 12 then 
			numberOfHoveddiagram = numberOfHoveddiagram + 1 
		end if	 
		
		'count diagrams named 'Hovediagram'
		if UCase(Mid((currentDiagram.Name),1,12)) = "HOVEDDIAGRAM" and len(currentDiagram.Name) > 12 then 
			numberOfHoveddiagramWithAdditionalInformationInTheName = numberOfHoveddiagramWithAdditionalInformationInTheName + 1 
		end if	 
	next
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Script Name: CheckOversiktsdiagram
' Author: Åsmund Tjora (based on FindHoveddiagramsInAS by Sara Henriksen)
' Date: 11.01.17
' Purpose: check if the applicationSchema-package has more than one diagram with a name starting with "Hoveddiagram", if so, check that there also is a
' diagram starting with "Oversiktsdiagram"
' /krav/oversiktsdiagram 
'@param[in]: package (EA.package) The package potentially containing diagrams with the provided name

sub CheckOversiktsdiagram(package)
	
	dim diagrams as EA.Collection
	set diagrams = package.Diagrams
	dim noHoveddiagram
	dim noOversiktsdiagram
	
	noHoveddiagram = 0
	noOversiktsdiagram = 0

	'find all diagrams in the package 
	dim i
	for i = 0 to diagrams.Count - 1
		dim currentDiagram as EA.Diagram
		set currentDiagram = diagrams.GetAt( i )
		if UCase(Mid(currentDiagram.Name,1,12)) = "HOVEDDIAGRAM" then 
			noHoveddiagram = noHoveddiagram + 1 
		end if	 
		if UCase(Mid(currentDiagram.Name,1,16)) = "OVERSIKTSDIAGRAM" then
			noOversiktsdiagram = noOversiktsdiagram + 1
		end if	 
	next
	if  ((noHoveddiagram > 1) and (noOversiktsdiagram = 0)) then
		session.output("Error: Package [" & package.Name & "] has more than one diagram with names starting with Hoveddiagram, but no diagram with name starting with Oversiktsdiagram [/krav/oversiktsdiagram]")
		globalErrorCounter = globalErrorCounter + 1 		
	end if
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Script Name: checkExternalCodelists
' Author: Sara Henriksen
' Date: 15.08.16
' Purpose: check each codeList for 'asDictionary' tag with value 'true', if so, check if tag codeList exist and if not return an error, if the value of the tag is empty also return an error
' /krav/eksternKodeliste
' 2 subs, 
'sub procedure to check if given codelist got the provided tag with value "true", if so, calls another sub procedure
'@param[in]: theElement (Attribute Class) and TaggedValueName (String)

sub checkExternalCodelists(theElement,  taggedValueName)

	if taggedValueName = "asDictionary" then 

		if not theElement is nothing and Len(taggedValueName) > 0 then

			'iterate trough all tagged values
			dim currentExistingTaggedValue AS EA.TaggedValue 
			dim taggedValuesCounter
			for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
				set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)

				'check if the tagged value exists 
				if currentExistingTaggedValue.Name = taggedValueName then
					'check if the value is "true" and if so, calls the subroutine to searching for codeList tags.
					if currentExistingTaggedValue.Value = "true" then 

						Call CheckCodelistTV(theElement, "codeList")
					end if 
				end if 
			next
		end if 
	end if 
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
'sub procedure to check if the provided tag exist (codeList), and if so, check  if the value is empty or not
'@param[in]: theElement (Element Class) and TaggedValueName (String)

sub CheckCodelistTV (theElement,  taggedValueNAME)

	'iterate tagged Values 
	dim currentExistingTaggedValue AS EA.TaggedValue 
	dim taggedValueCodeListMissing
	taggedValueCodeListMissing = true
	dim taggedValuesCounter
	
	for taggedValuesCounter = 0 to theElement.TaggedValues.Count - 1
		set currentExistingTaggedValue = theElement.TaggedValues.GetAt(taggedValuesCounter)
		'check if the tagged value exists
		if currentExistingTaggedValue.Name = taggedValueName then
			'Session.Output("følgende kodeliste:  " &theElement.Name)
			taggedValueCodeListMissing = false
			
			'if the codeList-value is empty, return an error 
			if currentExistingTaggedValue.Value = "" then 
				Session.Output("Error: Class [«"&theElement.Stereotype&"» "&theElement.Name& "] \ tag [codeList] lacks value. [/krav/eksternKodeliste]")
				globalErrorCounter = globalErrorCounter + 1 
			end if 
		end if 
	next
	
	'if the tagged value "codeList" is missing for an element(codelist), return an error
	if taggedValueCodeListMissing then
		Session.Output("Error: Class [«"&theElement.Stereotype&"» "&theElement.Name& "] lacks a [codeList] tag. [/krav/eksternKodeliste]")
		globalErrorCounter = globalErrorCounter + 1 
	end if
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


' -----------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: krav6-mnemoniskKodenavn
' Author: Kent Jonsrud
' Date: 2016-08-04
' Purpose: 
    'test if element name is legal NCName
    '/krav/6 - Navn på koder skal være mnemoniske (forståelige/huskbare), følge navnereglene for egenskapsnavn og være uten skilletegn og spesialtegn
    'Visuell sjekk om navnene er gode/forståelige - etter beste mnemoniske vurdering
    'Sjekk at navnet er NCName. 
	'Skilletegn og spesialtegn som må unngås er: blank, komma, !, "", #, $, %, &, ', (, ), *, +, /, :, ;, <, =, >, ?, @, [, \, ], ^, `, {, |, }, ~
	'((Tegnkoder under 32 (eksempelvis TAB) er ulovlige.))
    'Et modellelementnavn kan ikke starte med tall, ""-"" eller ""."""
	'Advarsel (Feil?) hvis kodens navn ikke er lowerCamelCase-NCName. 

sub krav6mnemoniskKodenavn(theElement)
	
	dim goodNames, lowerCameCase, badName
	goodNames = true
	lowerCameCase = true
	dim attr as EA.Attribute
	dim numberOfFaults
	numberOfFaults = 0
	dim numberOfWarnings
	numberOfWarnings = 0
	dim numberInList
	numberInList = 0
	
	'navigate through all attributes in the codeLists/enumeration 
	for each attr in theElement.Attributes
		'count number of attributes in one list
		numberInList = numberInList + 1 
		'check if the name is NCName
		if NOT IsNCName(attr.Name) then
			'count number of numeric initial values for one list
			numberOfFaults = numberOfFaults + 1
			Session.Output("Error: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] has illegal code name ["&attr.Name&"]. Recommended to use the script <lagLovligeNCNavnPåKodelistekoder>. [/krav/6]")
			if goodNames then
				badName = attr.Name
			end if
			goodNames = false 
		end if 
		'check if any of the names are lowerCameCase
		if NOT (mid(attr.Name,1,1) = LCASE(mid(attr.Name,1,1)) ) then
			numberOfWarnings = numberOfWarnings + 1
			if globalLogLevelIsWarning then
				Session.Output("Warning: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] has code name that is not lowerCamelCase ["&attr.Name&"]. Recommended to use the script <lagLovligeNCNavnPåKodelistekoder>. [/krav/6]")
			end if
			lowerCameCase = false
		End if
	next
	
	
	'if one or more names are illegal, return a error.
	if goodNames = false then 
		'Session.Output("Error: Illegal code names starts with ["&badName&"] for class: [«" &theElement.Stereotype& "» " &theElement.Name& "]. "&numberOfFaults&"/"&numberInList&" of the names are illegal.  Recommended to use the script <lagLovligeNCNavnPåKodelistekoder>   [/krav/6 ]")
		globalErrorCounter = globalErrorCounter +  numberOfFaults
	end if
	
	'if one or more names start with uppercase, return a warning.
	if lowerCameCase = false then 
		if globalLogLevelIsWarning then
			'Session.Output("Warning: All code names are not lowerCamelCase for class: [«" &theElement.Stereotype& "» " &theElement.Name& "].  Recommended to use the script <lagLovligeNCNavnPåKodelistekoder>  [/krav/6 ]")
			globalWarningCounter = globalWarningCounter +  numberOfWarnings
		end if	
	end if
end sub

Function IsNCName(streng)
    Dim txt, res, tegn, i, u
    u = true
	txt = ""
	For i = 1 To Len(streng)
        tegn = Mid(streng,i,1)
	    if tegn = " " or tegn = "," or tegn = """" or tegn = "#" or tegn = "$" or tegn = "%" or tegn = "&" or tegn = "(" or tegn = ")" or tegn = "*" Then
		    u=false
		end if 
	
		if tegn = "+" or tegn = "/" or tegn = ":" or tegn = ";" or tegn = "<" or tegn = ">" or tegn = "?" or tegn = "@" or tegn = "[" or tegn = "\" Then
		    u=false
		end if 
		If tegn = "]" or tegn = "^" or tegn = "`" or tegn = "{" or tegn = "|" or tegn = "}" or tegn = "~" or tegn = "'" or tegn = "´" or tegn = "¨" Then
		    u=false
		end if 
		if tegn <  " " then
		    u=false
		end if
	next
	tegn = Mid(streng,1,1)
	if tegn = "1" or tegn = "2" or tegn = "3" or tegn = "4" or tegn = "5" or tegn = "6" or tegn = "7" or tegn = "8" or tegn = "9" or tegn = "0" or tegn = "-" or tegn = "." Then
		u=false
	end if 
	IsNCName = u
End Function
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


' -----------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: krav7-kodedefinisjon
' Author: Kent Jonsrud
' Date: 2016-08-05
' Purpose: 
 	' test if element has definition
	'/krav/7
  	'Alle koder er konsepter, og skal ha tilstrekkelig definisjon. Det vil si alle unntatt lister over kjente egennavn.
  	'Visuell sjekk om navnene er egennavn, der dette ikke er tilfellet skal det finnes en definisjon
  	'Se Krav 3, bør kun gi advarsel fordi vi ikke kan sjekke om det dreier seg om et egetnavn eller ikke

sub krav7kodedefinisjon(theElement)
	
	dim goodNames, badName
	goodNames = true
	dim attr as EA.Attribute
	dim numberOfFaults
	numberOfFaults = 0
	dim numberInList
	numberInList = 0
	
	'navigate through all attributes in the codeLists/enumeration 
	for each attr in theElement.Attributes
		'count number of attributes in one list
		numberInList = numberInList + 1 
		'check if the code has definition
		if attr.Notes = "" then
			numberOfFaults = numberOfFaults + 1
			if globalLogLevelIsWarning then
				Session.Output("Warning: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] is missing definition for code ["&attr.Name&"]. [/krav/7]")
			end if
			if goodNames then
				badName = attr.Name
			end if
			goodNames = false 
		end if 
	next

	'if one or more codes lack definition, warning.
	if goodNames = false then 
		if globalLogLevelIsWarning then
			'Session.Output("Warning: Missing definition for code ["&badName&"] in class: [«" &theElement.Stereotype& "» " &theElement.Name& "]. "&numberOfFaults&"/"&numberInList&" of the codes lack definition. [/krav/7]")
			globalWarningCounter = globalWarningCounter + 1
		end if	
	end if
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


' -----------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: krav14 - inherit from same stereotype
' Author: Tore Johnsen
' Date: 2016-08-22
' Purpose: Checks that there is no inheritance between classes with unequal stereotypes.
'		/krav/14
' @param[in]: currentElement

sub krav14(currentElement)

	dim connectors as EA.Collection
	set connectors = currentElement.Connectors
	dim connectorsCounter
	
	for connectorsCounter = 0 to connectors.Count - 1
		dim currentConnector as EA.Connector
		set currentConnector = connectors.GetAt( connectorsCounter )
		dim targetElementID
		targetElementID = currentConnector.SupplierID
		dim elementOnOppositeSide as EA.Element
					
		if currentConnector.Type = "Generalization" then
			set elementOnOppositeSide = Repository.GetElementByID(targetElementID)
			
			if UCase(elementOnOppositeSide.Stereotype) <> UCase(currentElement.Stereotype) then
				session.output("Error: Class [" & elementOnOppositeSide.Name & "] has a stereotype that is not the same as the stereotype of [" & currentElement.Name & "]. A class can only inherit from a class with the same stereotype. [/krav/14]")
				globalErrorCounter = globalErrorCounter + 1 
			end if
		end if
	next
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


' -----------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: krav15-stereotyper
' Author: Kent Jonsrud
' Date: 2016-08-05
' Purpose: 
    '/krav/15
    'Modeller av geografisk informasjon skal ved behov bruke en av de standardiserte stereotypene, og ikke lage egne alternative stereotyper med samme mening.
    '(CodeList, dataType, enumeration, interface, Leaf, Union, FeatureType, ApplicationSchema) (Andre stereotyper med andre betydninger kan legges til.)
    'visuell sjekk at det ikke legges en annen betydning i stereotypene som er nevnt i kravet og  dersom stereotypen ikke er kjent.
    'Sjekk mot alle stereotyper som er nevnt i standarden og som er knyttet til et applikasjonsskjema.
    'Advarselsmelding der det er stereotyper som ikke er en del av lista.
    'NB - ta med<estimated>, beskrevet i ISO 19156 og SOSI Regler for UML modellering (testen skal være case-uavhengig)
    'høy
    'ta inn MessageType fra kap 9 i en senere versjon (2.0?)	Advarsel

sub krav15stereotyper(theElement)
	dim goodNames, badName, badStereotype, roleName
	goodNames = true
	dim attr as EA.Attribute
	dim conn as EA.Collection
	dim numberOfFaults
	numberOfFaults = 0
	dim numberInList
	numberInList = 0
	
	'navigate through all attributes  
	for each attr in theElement.Attributes
		numberInList = numberInList + 1 
		if attr.Stereotype <> "" then
			numberOfFaults = numberOfFaults + 1
			if globalLogLevelIsWarning then
				Session.Output("Warning: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] has unknown stereotype. «" & attr.Stereotype & "» on attribute ["&attr.Name&"]. [/krav/15]")
				globalWarningCounter = globalWarningCounter + 1
			end if	
			if goodNames then
				badName = attr.Name
				badStereotype = attr.Stereotype
			end if
			goodNames = false 
		end if 
	next
	
	'if one or more codes lack definition, warning.
	if goodNames = false then 
		if globalLogLevelIsWarning then
			'Session.Output("Warning: Unknown attribute stereotypes starting with [«"&badStereotype&"» "&badName&"] in class: [«" &theElement.Stereotype& "» " &theElement.Name& "]. "&numberOfFaults&"/"&numberInList&" of the attributes have unknown stereotype. [/krav/15]")
			globalWarningCounter = globalWarningCounter + 1
		end if	
	end if

	'operations?
	
	'Association roles with stereotypes other than «estimated»
	for each conn in theElement.Connectors
		roleName = ""
		badStereotype = ""
		if theElement.ElementID = conn.ClientID then
			roleName = conn.SupplierEnd.Role
			badStereotype = conn.SupplierEnd.Stereotype
		end if
		if theElement.ElementID = conn.SupplierID then
			roleName = conn.ClientEnd.Role
			badStereotype = conn.ClientEnd.Stereotype
		end if
		'(ignoring all association roles without name!)
		if roleName <> "" then
			if badStereotype <> "" and LCase(badStereotype) <> "estimated" then
				if globalLogLevelIsWarning then
					Session.Output("Warning: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] as unknown stereotype «"&badStereotype&"» on role name ["&roleName&"]. [/krav/15]")				
					globalWarningCounter = globalWarningCounter + 1 
				end if	
			end if
		end if
	next
	
	'Associations with stereotype, especially «topo»
	for each conn in theElement.Connectors
		if conn.Stereotype <> "" then
			if LCase(conn.Stereotype) = "topo" then
 				Session.Output("Error: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] has illegal stereotype «"&conn.Stereotype&"» on association named ["&conn.Name&"]. Recommended to use the script <endreTopoAssosiasjonTilRestriksjon>. [/krav/15]")				
 				globalErrorCounter = globalErrorCounter + 1 
			else
				if globalLogLevelIsWarning then
					Session.Output("Warning: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] has unknown stereotype «"&conn.Stereotype&"» on association named ["&conn.Name&"]. [/krav/15]")				
					globalWarningCounter = globalWarningCounter + 1 
				end if	
			end if
		end if
	next
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


' -----------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: krav16-unikeNCnavn
' Author: Kent Jonsrud
' Date: 2016-08-09
' Purpose: 
    '/krav/16
    'Alle navn på modellelementer skal være case-insensitivt unike innenfor sitt navnerom, og ikke inneholde blanke eller andre skilletegn.
    'Merknad: navnerommet til roller og egenskaper er klassen.
    'Sjekk at navnene til klasser (classifier: kodelister, enumerations, datatyper, objekttyper) og underpakker(!), er unike innenfor sitt navnerom (valgt pakke)
    'Navn til roller, egenskaper og operasjoner skal være unike innenfor klassen.
    'Notat: NCName, unike navn på klasse i underpakker, unike eg-/rolle-/oper-navn (forby polymorfisme på operasjoner?)
 
sub krav16unikeNCnavn(theElement)
	
	dim goodNames, lowerCameCase, badName, roleName
	goodNames = true
	lowerCameCase = true
	dim super as EA.Element
	dim attr as EA.Attribute
	dim oper as EA.Collection
	dim conn as EA.Collection
	dim numberOfFaults
	numberOfFaults = 0
	dim numberInList
	numberInList = 0

	dim PropertyNames
	Set PropertyNames = CreateObject("System.Collections.ArrayList")

	'List of element IDs to check for endless recursion (Åsmund)
	dim inheritanceElementList
	set inheritanceElementList = CreateObject("System.Collections.ArrayList")

	'Association role names
	for each conn in theElement.Connectors
		roleName = ""
		if theElement.ElementID = conn.ClientID then
			roleName = conn.SupplierEnd.Role
		end if
		if theElement.ElementID = conn.SupplierID then
			roleName = conn.ClientEnd.Role
		end if
		'(ignoring all association roles without name!)
		if roleName <> "" then
			if PropertyNames.IndexOf(UCase(roleName),0) <> -1 then
				Session.Output("Error: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] has non-unique role name ["&roleName&"]. [/krav/16]")				
 				globalErrorCounter = globalErrorCounter + 1 
			else
				PropertyNames.Add UCase(roleName)
			end if
			if NOT IsNCName(roleName) then
				Session.Output("Error: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] has illegal role name, ["&roleName&"] is not a NCName. [/krav/16]")				
 				globalErrorCounter = globalErrorCounter + 1 
			end if
		end if
	next
	
	'Operation names
	for each oper in theElement.Methods
		if PropertyNames.IndexOf(UCase(oper.Name),0) <> -1 then
			Session.Output("Error: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] has non-unique operation property name ["&oper.Name&"]. [/krav/16]")				
			globalErrorCounter = globalErrorCounter + 1 
		else
			PropertyNames.Add UCase(oper.Name)
		end if
		'check if the name is NCName
		if NOT IsNCName(oper.Name) then
				Session.Output("Error: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] has illegal operation name, ["&oper.Name&"] is not a NCName. [/krav/16]")				
 				globalErrorCounter = globalErrorCounter + 1 
		end if 
	next
	
	'Constraint names TODO
	
	'navigate through all attributes 
	for each attr in theElement.Attributes
		'count number of attributes in one list
		numberInList = numberInList + 1 
		if PropertyNames.IndexOf(UCase(attr.Name),0) <> -1 then
			Session.Output("Error: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] has non-unique attribute property name ["&attr.Name&"]. [/krav/16]")				
			globalErrorCounter = globalErrorCounter + 1 
		else
			PropertyNames.Add UCase(attr.Name)
		end if

		'check if the name is NCName (exception for code names - they have a separate test.)
		if NOT ((theElement.Type = "Class") and (UCase(theElement.Stereotype) = "CODELIST"  Or UCase(theElement.Stereotype) = "ENUMERATION")) then
			if NOT IsNCName(attr.Name) then
				'count number of numeric initial values for one list
				Session.Output("Error: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] has illegal attribute name, ["&attr.Name&"] is not a NCName. [/krav/16]")				
 				globalErrorCounter = globalErrorCounter + 1 
			end if
		end if 
	next

	'Other attributes and roles inherited from outside package
	'Traverse and test against inherited names but do not add the inherited names to the list(!)
	for each conn in theElement.Connectors

		if conn.Type = "Generalization" then
			if theElement.ElementID = conn.ClientID then
				set super = Repository.GetElementByID(conn.SupplierID)
				
				'Check agains endless recursion (Åsmund)
				dim hopOutOfEndlessRecursion
				dim inheritanceElementID
				hopOutOfEndlessRecursion = 0
				inheritanceElementList.Add(theElement.ElementID)
				for each inheritanceElementID in inheritanceElementList
					if inheritanceElementID = super.ElementID then 
						hopOutOfEndlessRecursion = 1
						Session.Output("Error: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] is a generalization of itself.")
						globalErrorCounter = globalErrorCounter + 1
					end if
				next
				if hopOutOfEndlessRecursion=0 then call krav16unikeNCnavnArvede(super, PropertyNames, inheritanceElementList)
			end if
		end if
	next
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


' -----------------------------------------------------------START-------------------------------------------------------------------------------------------
sub krav16unikeNCnavnArvede(theElement, PropertyNames, inheritanceElementList)
	dim goodNames, lowerCameCase, badName, roleName
	goodNames = true
	lowerCameCase = true
	dim super as EA.Element
	dim attr as EA.Attribute
	dim oper as EA.Collection
	dim conn as EA.Collection
 	dim numberOfFaults
	numberOfFaults = 0
	dim numberInList
	numberInList = 0

'	test if supertype name is same as one in the tested package. (supertype may well be outside the tested package.)
'	if ClassAndPackageNames.IndexOf(UCase(theElement.Name),0) <> -1 then
'	Session.Output("Warning: non-unique supertype name [«" &theElement.Stereotype& "» "&theElement.Name&"] in package: ["&Repository.GetPackageByID(theElement.PackageID).Name&"].  EA-type:" &theElement.Type& "  [/krav/16 ]")				
' 	globalWarningCounter = globalWarningCounter + 1
'	end if

	'Association role names
	for each conn in theElement.Connectors

		roleName = ""
		if theElement.ElementID = conn.ClientID then
			roleName = conn.SupplierEnd.Role
		end if
		if theElement.ElementID = conn.SupplierID then
			roleName = conn.ClientEnd.Role
		end if
		'(ignoring all association roles without name!)
		if roleName <> "" then
			if PropertyNames.IndexOf(UCase(roleName),0) <> -1 then
				if globalLogLevelIsWarning then
					Session.Output("Warning: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] in package: ["&Repository.GetPackageByID(theElement.PackageID).Name&"] has non-unique inherited role property name ["&roleName&"] implicitly redefined from. [/krav/16]")				
					globalWarningCounter = globalWarningCounter + 1
				end if	
			end if
		end if
	next
	
	'Operation names
	for each oper in theElement.Methods
		if PropertyNames.IndexOf(UCase(oper.Name),0) <> -1 then
			if globalLogLevelIsWarning then
				Session.Output("Warning: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] in package: ["&Repository.GetPackageByID(theElement.PackageID).Name&"] has inherited and implicitly redefined non-unique operation property name ["&oper.Name&"]. [/krav/16]")				
				globalWarningCounter = globalWarningCounter + 1
			end if	
		end if
	next
	
	'Constraint names TODO
	
	'navigate through all attributes 
	for each attr in theElement.Attributes
		'count number of attributes in one list
		numberInList = numberInList + 1 
		if PropertyNames.IndexOf(UCase(attr.Name),0) <> -1 then
			if globalLogLevelIsWarning then
				Session.Output("Warning: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] in package: ["&Repository.GetPackageByID(theElement.PackageID).Name&"] has non-unique inherited and implicitly redefined attribute property name["&attr.Name&"]. [/krav/16]")				
				globalWarningCounter = globalWarningCounter + 1
			end if	
		end if
	next

	'Other attributes and roles inherited from outside package
	'Traverse and test against inherited names but do not add the inherited names to the list
	for each conn in theElement.Connectors
		if conn.Type = "Generalization" then
			if theElement.ElementID = conn.ClientID then
				set super = Repository.GetElementByID(conn.SupplierID)
				'Check agains endless recursion (Åsmund)
				dim hopOutOfEndlessRecursion
				dim inheritanceElementID
				hopOutOfEndlessRecursion = 0
				inheritanceElementList.Add(theElement.ElementID)
				for each inheritanceElementID in inheritanceElementList
					if inheritanceElementID = super.ElementID then 
						hopOutOfEndlessRecursion = 1
						Session.Output("Error: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] is a generalization of itself.")
						globalErrorCounter = globalErrorCounter + 1
					end if
				next
				if hopOutOfEndlessRecursion=0 then call krav16unikeNCnavnArvede(super, PropertyNames, inheritanceElementList)
			end if
		end if
	next
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


' -----------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: reqUmlProfile
' Author: Kent Jonsrud
' Date: 2016-08-08
' Purpose: 
    '/req/uml/profile     ~ bygger på /krav/22 og /krav/25
    'Applikasjonsskjema skal modelleres ved bruk av UML-profilen definert i ISO19103:2015, og med tillegg beskrevet i dette kapittel. (Kapittel 11 i SOSI regler for UML-modellering 5.0)


sub reqUmlProfile(theElement)
	
	dim attr as EA.Attribute

	'List of well known core and extension type names defined in iso 19103:2015
	dim ExtensionTypes
	Set ExtensionTypes = CreateObject("System.Collections.ArrayList")
	ExtensionTypes.Add "Date"
	ExtensionTypes.Add "Time"
	ExtensionTypes.Add "DateTime"
	ExtensionTypes.Add "CharacterString"
	ExtensionTypes.Add "Number"
	ExtensionTypes.Add "Decimal"
	ExtensionTypes.Add "Integer"
	ExtensionTypes.Add "Real"
	ExtensionTypes.Add "Boolean"
	ExtensionTypes.Add "Vector"

	ExtensionTypes.Add "Bit"
	ExtensionTypes.Add "Digit"
	ExtensionTypes.Add "Sign"

	ExtensionTypes.Add "NameSpace"
	ExtensionTypes.Add "GenericName"
	ExtensionTypes.Add "LocalName"
	ExtensionTypes.Add "ScopedName"
	ExtensionTypes.Add "TypeName"
	ExtensionTypes.Add "MemberName"

	ExtensionTypes.Add "Any"

	ExtensionTypes.Add "Record"
	ExtensionTypes.Add "RecordType"
	ExtensionTypes.Add "Field"
	ExtensionTypes.Add "FieldType"
	
	'iso 19103 Annex-C types
	ExtensionTypes.Add "LanguageString"
	
	ExtensionTypes.Add "Anchor"
	ExtensionTypes.Add "FileName"
	ExtensionTypes.Add "MediaType"
	ExtensionTypes.Add "URI"
	
	ExtensionTypes.Add "UnitOfMeasure"
	ExtensionTypes.Add "UomArea"
	ExtensionTypes.Add "UomLenght"
	ExtensionTypes.Add "UomAngle"
	ExtensionTypes.Add "UomAcceleration"
	ExtensionTypes.Add "UomAngularAcceleration"
	ExtensionTypes.Add "UomAngularSpeed"
	ExtensionTypes.Add "UomSpeed"
	ExtensionTypes.Add "UomCurrency"
	ExtensionTypes.Add "UomVolume"
	ExtensionTypes.Add "UomTime"
	ExtensionTypes.Add "UomScale"
	ExtensionTypes.Add "UomWeight"
	ExtensionTypes.Add "UomVelocity"

	ExtensionTypes.Add "Measure"
	ExtensionTypes.Add "Length"
	ExtensionTypes.Add "Distance"
	ExtensionTypes.Add "Speed"
	ExtensionTypes.Add "Angle"
	ExtensionTypes.Add "Scale"
	ExtensionTypes.Add "TimeMeasure"
	ExtensionTypes.Add "Area"
	ExtensionTypes.Add "Volume"
	ExtensionTypes.Add "Currency"
	ExtensionTypes.Add "Weight"
	ExtensionTypes.Add "AngularSpeed"
	ExtensionTypes.Add "DirectedMeasure"
	ExtensionTypes.Add "Velocity"
	ExtensionTypes.Add "AngularVelocity"
	ExtensionTypes.Add "Acceleration"
	ExtensionTypes.Add "AngularAcceleration"
	
	'well known and often used spatial types from iso 19107:2003
	ExtensionTypes.Add "DirectPosition"
	ExtensionTypes.Add "GM_Object"
	ExtensionTypes.Add "GM_Primitive"
	ExtensionTypes.Add "GM_Complex"
	ExtensionTypes.Add "GM_Aggregate"
	ExtensionTypes.Add "GM_Point"
	ExtensionTypes.Add "GM_Curve"
	ExtensionTypes.Add "GM_Surface"
	ExtensionTypes.Add "GM_Solid"
	ExtensionTypes.Add "GM_MultiPoint"
	ExtensionTypes.Add "GM_MultiCurve"
	ExtensionTypes.Add "GM_MultiSurface"
	ExtensionTypes.Add "GM_MultiSolid"
	ExtensionTypes.Add "GM_CompositePoint"
	ExtensionTypes.Add "GM_CompositeCurve"
	ExtensionTypes.Add "GM_CompositeSurface"
	ExtensionTypes.Add "GM_CompositeSolid"
	ExtensionTypes.Add "TP_Object"
	'ExtensionTypes.Add "TP_Primitive"
	ExtensionTypes.Add "TP_Complex"
	ExtensionTypes.Add "TP_Node"
	ExtensionTypes.Add "TP_Edge"
	ExtensionTypes.Add "TP_Face"
	ExtensionTypes.Add "TP_Solid"
	ExtensionTypes.Add "TP_DirectedNode"
	ExtensionTypes.Add "TP_DirectedEdge"
	ExtensionTypes.Add "TP_DirectedFace"
	ExtensionTypes.Add "TP_DirectedSolid"
	ExtensionTypes.Add "GM_OrientableCurve"
	ExtensionTypes.Add "GM_OrientableSurface"
	ExtensionTypes.Add "GM_PolyhedralSurface"
	ExtensionTypes.Add "GM_triangulatedSurface"
	ExtensionTypes.Add "GM_Tin"

	'well known and often used coverage types from iso 19123:2007
	ExtensionTypes.Add "CV_Coverage"
	ExtensionTypes.Add "CV_DiscreteCoverage"
	ExtensionTypes.Add "CV_DiscretePointCoverage"
	ExtensionTypes.Add "CV_DiscreteGridPointCoverage"
	ExtensionTypes.Add "CV_DiscreteCurveCoverage"
	ExtensionTypes.Add "CV_DiscreteSurfaceCoverage"
	ExtensionTypes.Add "CV_DiscreteSolidCoverage"
	ExtensionTypes.Add "CV_ContinousCoverage"
	ExtensionTypes.Add "CV_ThiessenPolygonCoverage"
	'ExtensionTypes.Add "CV_ContinousQuadrilateralGridCoverageCoverage"
	ExtensionTypes.Add "CV_ContinousQuadrilateralGridCoverage"
	ExtensionTypes.Add "CV_HexagonalGridCoverage"
	ExtensionTypes.Add "CV_TINCoverage"
	ExtensionTypes.Add "CV_SegmentedCurveCoverage"

	'well known and often used temporal types from iso 19108:2006/2002?
	ExtensionTypes.Add "TM_Instant"
	ExtensionTypes.Add "TM_Period"
	ExtensionTypes.Add "TM_Node"
	ExtensionTypes.Add "TM_Edge"
	ExtensionTypes.Add "TM_TopologicalComplex"
	
	'well known and often used observation related types from OM_Observation in iso 19156:2011
	ExtensionTypes.Add "TM_Object"
	ExtensionTypes.Add "DQ_Element"
	ExtensionTypes.Add "NamedValue"
	
	'well known and often used quality element types from iso 19157:2013
	ExtensionTypes.Add "DQ_AbsoluteExternalPositionalAccurracy"
	ExtensionTypes.Add "DQ_RelativeInternalPositionalAccuracy"
	ExtensionTypes.Add "DQ_AccuracyOfATimeMeasurement"
	ExtensionTypes.Add "DQ_TemporalConsistency"
	ExtensionTypes.Add "DQ_TemporalValidity"
	ExtensionTypes.Add "DQ_ThematicClassificationCorrectness"
	ExtensionTypes.Add "DQ_NonQuantitativeAttributeCorrectness"
	ExtensionTypes.Add "DQ_QuanatitativeAttributeAccuracy"

	'well known and often used metadata element types from iso 19115-1:200x and iso 19139:2x00x
	ExtensionTypes.Add "PT_FreeText"
	ExtensionTypes.Add "LocalisedCharacterString"
	ExtensionTypes.Add "MD_Resolution"
	'ExtensionTypes.Add "CI_Citation"
	'ExtensionTypes.Add "CI_Date"

	'other less known Norwegian geometry types
	ExtensionTypes.Add "Punkt"
	ExtensionTypes.Add "Kurve"
	ExtensionTypes.Add "Flate"
	ExtensionTypes.Add "Sverm"

	'navigate through all attributes 
	for each attr in theElement.Attributes
		'count number of attributes in one list
		if attr.ClassifierID = 0 then
			'check if the attribute has a well known core type
			if ExtensionTypes.IndexOf(attr.Type,0) = -1 then	
				Session.Output("Error: Class [«" &theElement.Stereotype& "» " &theElement.Name& "] has unknown type for attribute ["&attr.Name&" : "&attr.Type&"]. [/req/uml/profile]")
				globalErrorCounter = globalErrorCounter + 1 
			end if
		end if 
	next
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: krav18-viseAlt
' Author: Kent Jonsrud
' Date: 2016-08-09..30, 2016-09-05, 2017-01-17 (no more false positives)
' Purpose: test whether a class is showing all its content in at least one class diagram.
    '/krav/18

sub krav18viseAlt(theElement)

	dim diagram as EA.Diagram
	dim diagrams as EA.Collection
	dim diao as EA.DiagramObject
	dim conn as EA.Collection
	dim super as EA.Element
	dim base as EA.Collection
	dim child as EA.Collection
	dim embed as EA.Collection
	dim realiz as EA.Collection
	dim viserAlt
	viserAlt = false
	
	'navigate through all diagrams and find those the element knows
	Dim i, shownTimes
	shownTimes=0
	For i = 0 To diaoList.Count - 1
		if theElement.ElementID = diaoList.GetByIndex(i) then
			set diagram = Repository.GetDiagramByID(diagList.GetByIndex(i))
			shownTimes = shownTimes + 1
			for each diao in diagram.DiagramObjects
				if diao.ElementID = theElement.ElementID then
					exit for
				end if
			next

			if theElement.Attributes.Count = 0 or InStr(1,diagram.ExtendedStyle,"HideAtts=1") = 0 then
				if theElement.Methods.Count = 0 or InStr(1,diagram.ExtendedStyle,"HideOps=1") = 0 then
					if InStr(1,diagram.ExtendedStyle,"HideEStereo=1") = 0 then
						if InStr(1,diagram.ExtendedStyle,"UseAlias=1") = 0 or theElement.Alias = "" then
							if (showAllProperties(theElement, diagram, diao)) then
								'shows all OK in this diagram, how about inherited?
								viserAlt = true
							end if
						end if
					end if
				end if
			end if

		end if
	next
	
	if NOT viserAlt then
 		globalErrorCounter = globalErrorCounter + 1 
 		if shownTimes = 0 then
			Session.Output("Error: Class [«" &theElement.Stereotype& "» "&theElement.Name&"] is not shown in any diagram. [/krav/18]")
		else
			Session.Output("Error: Class [«" &theElement.Stereotype& "» "&theElement.Name&"] is not shown fully in at least one diagram. [/krav/18]")				
		end if
	end if
end sub

function showAllProperties(theElement, diagram, diao)
	showAllProperties = false
	if InStr(1,diagram.ExtendedStyle,"HideAtts=1") = 0 and diao.ShowPublicAttributes or InStr(1,diao.Style,"AttCustom=0" ) <> 0 or theElement.Attributes.Count = 0 then
		if InStr(1,diagram.ExtendedStyle,"HideOps=1") = 0 and diao.ShowPublicOperations or InStr(1,diao.Style,"OpCustom=0" ) <> 0 or theElement.Methods.Count = 0 then
			if InStr(1,diagram.ExtendedStyle,"ShowCons=0") = 0 or diao.ShowConstraints or InStr(1,diao.Style,"Constraint=1" ) <> 0 or theElement.Constraints.Count = 0 then
				' all attribute parts really shown? ...
				if InStr(1,diagram.StyleEX,"VisibleAttributeDetail=1" ) = 0 or theElement.Attributes.Count = 0 then
					showAllProperties = true
				end if
			end if
		end if
	end if
end function




'Recursive loop through subpackages, creating a list of all model elements and their corresponding diagrams
sub recListDiagramObjects(p)
	
	dim d as EA.Diagram
	dim Dobj as EA.DiagramObject
	for each d In p.diagrams
		for each Dobj in d.DiagramObjects
			If not diaoList.ContainsKey(Dobj.ElementID) Then
				diaoList.Add Dobj.InstanceID, Dobj.ElementID
				diagList.Add Dobj.InstanceID, Dobj.DiagramID
			end if   
		next	
	next
		
	dim subP as EA.Package
	for each subP in p.packages
	    recListDiagramObjects(subP)
	next
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub name: krav12
' Author: Magnus Karge
' Date: 20170110 
' Purpose:  sub procedure to check if a given dataType element's (element with stereotype DataType or of type DataType) associations are 
'			compositions and the composition is on the correct end (datatypes must only be targets of compositions)
' 			Implementation of /krav/navning
' 			
' @param[in]: 	theElement (EA.Element). The element to check. Can only be classifier of type data type or with stereotype dataType
'				theConnector (EA.Connector). The connector/association between theElement and theElementOnOppositeSide
'				theElementOnOppositeSide (EA.Element). The classifier on the other side of the connector/association
 
sub krav12(theElement, theConnector, theElementOnOppositeSide)
	dim currentElement AS EA.Element
	set currentElement = theElement
	dim elementOnOppositeSide AS EA.Element
	set elementOnOppositeSide = theElementOnOppositeSide
	dim currentConnector AS EA.Connector
	set currentConnector = theConnector
	
	dim dataTypeOnBothSides
	if (Ucase(currentElement.Stereotype) = Ucase("dataType") or currentElement.Type = "DataType") and (Ucase(elementOnOppositeSide.Stereotype) = Ucase("dataType") or elementOnOppositeSide.Type = "DataType") then
		dataTypeOnBothSides = true
	else	
		dataTypeOnBothSides = false
	end if
								
	'check if the elementOnOppositeSide has stereotype "dataType" and this side's end is no composition and not elements both sides of the association are datatypes
	if (Ucase(elementOnOppositeSide.Stereotype) = Ucase("dataType")) and not (currentConnector.ClientEnd.Aggregation = 2) and not dataTypeOnBothSides then 
		Session.Output( "Error: Class [«"&elementOnOppositeSide.Stereotype&"» "& elementOnOppositeSide.Name &"] has association to class [" & currentElement.Name & "] that is not a composition on "& currentElement.Name &"-side. [/krav/12]")									 
		globalErrorCounter = globalErrorCounter + 1 
	end if 

	'check if this side's element has stereotype "dataType" and the opposite side's end is no composition 
	if (Ucase(currentElement.Stereotype) = Ucase("dataType")) and not (currentConnector.SupplierEnd.Aggregation = 2) and not dataTypeOnBothSides then 
		Session.Output( "Error: Class [«"&currentElement.Stereotype&"» "& currentElement.Name &"] has association to class [" & elementOnOppositeSide.Name & "] that is not a composition on "& elementOnOppositeSide.Name &"-side. [/krav/12]")									 
		globalErrorCounter = globalErrorCounter + 1 
	end if 

end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub name: krav10
' Author: Magnus Karge
' Date: 20170110 
' Purpose:  sub procedure to check if the given association properties fulfill the requirements regarding
'			multiplicity on navigable ends (navigable ends shall have multiplicity)
' 			
' @param[in]: 	theElement (EA.Element). The element that "ownes" the association to check
'				sourceEndNavigable (CharacterString). navigable setting on association's source end
'				targetEndNavigable (CharacterString). navigable setting on association's target end
'				sourceEndName (CharacterString). role name on association's source end
'				targetEndName (CharacterString). role name on association's target end
'				sourceEndCardinality (CharacterString). multiplicity on association's source end
'				targetEndCardinality (CharacterString). multiplicity on association's target end
sub krav10(theElement, sourceEndNavigable, targetEndNavigable, sourceEndName, targetEndName, sourceEndCardinality, targetEndCardinality)
	if sourceEndNavigable = "Navigable" and sourceEndCardinality = "" then 
		Session.Output( "Error: Class [«"&theElement.Stereotype&"» "& theElement.Name &"] \ association role [" & sourceEndName & "] lacks multiplicity. [/krav/10]") 
		globalErrorCounter = globalErrorCounter + 1 
	end if 
 								 
	if targetEndNavigable = "Navigable" and targetEndCardinality = "" then 
		Session.Output( "Error: Class [«"&theElement.Stereotype&"» "& theElement.Name &"] \ association role [" & targetEndName & "] lacks multiplicity. [/krav/10]") 
		globalErrorCounter = globalErrorCounter + 1 
	end if 
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub name: krav11
' Author: Magnus Karge
' Date: 20170110 
' Purpose:  sub procedure to check if the given association has role names on navigable ends 
'			(navigable ends shall have role names)
' 			
' @param[in]: 	theElement (EA.Element). The element that "ownes" the association to check
'				sourceEndNavigable (CharacterString). navigable setting on association's source end
'				targetEndNavigable (CharacterString). navigable setting on association's target end
'				sourceEndName (CharacterString). role name on association's source end
'				targetEndName (CharacterString). role name on association's target end
'				elementOnOppositeSide (EA.Element). The element on the opposite side of the association to check
sub krav11(theElement, sourceEndNavigable, targetEndNavigable, sourceEndName, targetEndName, elementOnOppositeSide)
	if sourceEndNavigable = "Navigable" and sourceEndName = "" then 
		Session.Output( "Error: Association between class [«"&theElement.Stereotype&"» "& theElement.Name &"] and class [«"&elementOnOppositeSide.Stereotype&"» "& elementOnOppositeSide.Name & "] lacks role name on navigable end on "& theElement.Name &"-side. [/krav/11]") 
		globalErrorCounter = globalErrorCounter + 1 
	end if 
 								 
	if targetEndNavigable = "Navigable" and targetEndName = "" then 
		Session.Output( "Error: Association between class [«"&theElement.Stereotype&"» "& theElement.Name &"] and class [«"&elementOnOppositeSide.Stereotype&"» "& elementOnOppositeSide.Name & "] lacks role name on navigable end on "& elementOnOppositeSide.Name &"-side. [/krav/11]") 
		globalErrorCounter = globalErrorCounter + 1 
	end if 
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub name: checkRoleNames
' Author: Magnus Karge
' Date: 20170110 
' Purpose:  sub procedure to check if a given association's role names start with lower case 
'			(navigable ends shall have role names [krav/navning]) 
' 			
' @param[in]: 	theElement (EA.Element). The element that "ownes" the association to check
'				sourceEndName (CharacterString). role name on association's source end
'				targetEndName (CharacterString). role name on association's target end
'				elementOnOppositeSide (EA.Element). The element on the opposite side of the association to check
sub checkRoleNames(theElement, sourceEndName, targetEndName, elementOnOppositeSide)
	if not sourceEndName = "" and not Left(sourceEndName,1) = LCase(Left(sourceEndName,1)) then 
		Session.Output("Error: Role name [" & sourceEndName & "] on association end connected to class ["& theElement.Name &"] shall start with lowercase letter. [/krav/navning]") 
		globalErrorCounter = globalErrorCounter + 1 
	end if 

	if not (targetEndName = "") and not (Left(targetEndName,1) = LCase(Left(targetEndName,1))) then 
		Session.Output("Error: Role name [" & targetEndName & "] on association end connected to class ["& elementOnOppositeSide.Name &"] shall start with lowercase letter. [/krav/navning]") 
		globalErrorCounter = globalErrorCounter + 1 
	end if 
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Script Name: checkEndingOfPackageName
' Author: Sara Henriksen, Åsmund Tjora	
' Purpose: check if the package name ends with a version number. The version number could be a date or a serial number. Returns an error if the version 
' number contains anything other than 0-2 dots or numbers. 
' Packages under development should have the text "Utkast" as the final element, after the version number. 
' Date: 25.08.16 (original version) 10.01.17 (Updated version)
sub checkEndingOfPackageName(thePackage)
	if UCase(thePackage.Element.Stereotype)="APPLICATIONSCHEMA" then
		'find the last part of the package name, after "-" 
		dim startContent, endContent, stringContent, cleanContent 		
		
		'remove any "Utkast" part of the name 
		cleanContent=replace(UCase(thePackage.Name), "UTKAST", "")
		
		endContent = len(cleanContent)
	
		startContent = InStr(cleanContent, "-") 
	
		stringContent = mid(cleanContent, startContent+1, endContent) 	
		dim versionNumberInPackageName
		versionNumberInPackageName = false 
		'count number of dots, only allowed to use max two. 
		dim dotCounter
		dotCounter = 0

		'check that the package name contains a "-", and thats it is just number(s) and "." after. 
		if InStr(thePackage.Name, "-") then 			
			'if the string is numeric or it has dots, set the valueOk true 
			if  InStr(stringContent, ".")  or IsNumeric(stringContent)  then
				versionNumberInPackageName = true 
				dim i, tegn 
				for i = 1 to len(stringContent) 
					tegn = Mid(stringContent,i,1)
					if tegn = "." then
						dotCounter = dotCounter  + 1 
					end if 
				next 
				'count number of dots. If it's more than 2 return an error. 
				if dotCounter < 3 then 
					versionNumberInPackageName = true
				else 
					versionNumberInPackageName = false
				end if
			end if 
		end if 

		'check the string for letters and symbols. If the package name contains one of the following, then return an error. 
		if inStr(UCase(stringContent), "A") or inStr(UCase(stringContent), "B") or inStr(UCase(stringContent), "C") or inStr(UCase(stringContent), "D") or inStr(UCase(stringContent), "E") or inStr(UCase(stringContent), "F") or inStr(UCase(stringContent), "G") or inStr(UCase(stringContent), "H") or inStr(UCase(stringContent), "I") or inStr(UCase(stringContent), "J") or inStr(UCase(stringContent), "K") or inStr(UCase(stringContent), "L")  then 
			versionNumberInPackageName = false
		end if 	
		if inStr(UCase(stringContent), "M") or inStr(UCase(stringContent), "N") or inStr(UCase(stringContent), "O") or inStr(UCase(stringContent), "P") or inStr(UCase(stringContent), "Q") or inStr(UCase(stringContent), "R") or inStr(UCase(stringContent), "S") or inStr(UCase(stringContent), "T") or inStr(UCase(stringContent), "U") or inStr(UCase(stringContent), "V") or inStr(UCase(stringContent), "W") or inStr(UCase(stringContent), "X") then          
			versionNumberInPackageName = false
		end if 
		if inStr(UCase(stringContent), "Y") or inStr(UCase(stringContent), "Z") or inStr(UCase(stringContent), "Æ") or inStr(UCase(stringContent), "Ø") or inStr(UCase(stringContent), "Å") then 
			versionNumberInPackageName = false
		end if 
		if inStr(stringContent, ",") or inStr(stringContent, "!") or inStr(stringContent, "@") or inStr(stringContent, "%") or inStr(stringContent, "&") or inStr(stringContent, """") or inStr(stringContent, "#") or inStr(stringContent, "$") or inStr(stringContent, "'") or inStr(stringContent, "(") or inStr(stringContent, ")") or inStr(stringContent, "*") or inStr(stringContent, "+") or inStr(stringContent, "/") then        
			versionNumberInPackageName = false
		end if
		if inStr(stringContent, ":") or inStr(stringContent, ";") or inStr(stringContent, ">") or inStr(stringContent, "<") or inStr(stringContent, "=") then
			versionNumberInPackageName = false
		end if 
	
		if versionNumberInPackageName = false  then  
			Session.Output("Error: Package ["&thePackage.Name&"] does not have a name ending with a version number. [/krav/SOSI-modellregister/applikasjonsskjema/versjonsnummer]")
			globalErrorCounter = globalErrorCounter + 1	
		end if 
	end if	
end sub 
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------

'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub name: checkUniqueFeatureTypeNames
' Author: Magnus Karge
' Date: 20170110 
' Purpose:  sub procedure to check if a given FeatureType's name is unique within the applicationSchema
''			(the class name shall be unique within the application schema [/req/uml/feature]) 
' 			
' @param[in]: 	none - uses only global variables FeatureTypeNames and FeatureTypeElementIDs
sub checkUniqueFeatureTypeNames()
	'iterate over elements in the  name and id arrays until the arrays are empty
	DO UNTIL FeatureTypeNames.count = 0 AND FeatureTypeElementIDs.count = 0 				
		dim temporaryFeatureTypeArray
		set temporaryFeatureTypeArray = CreateObject("System.Collections.ArrayList")
		dim ftNameToCompare
		ftNameToCompare = FeatureTypeNames.Item(0)
		dim ftElementID
		ftElementID = FeatureTypeElementIDs.Item(0)
		dim initialElementToAdd AS EA.Element
		set initialElementToAdd = Repository.GetElementByID(ftElementID)
		temporaryFeatureTypeArray.Add(initialElementToAdd)
		FeatureTypeNames.RemoveAt(0)
		FeatureTypeElementIDs.RemoveAt(0)
		dim elementNumber
		for elementNumber = FeatureTypeNames.count - 1 to 0 step -1
			dim currentName
			currentName = FeatureTypeNames.Item(elementNumber)
			if currentName = ftNameToCompare then
				dim currentElementID
				currentElementID = FeatureTypeElementIDs.Item(elementNumber)
				dim additionalElementToAdd AS EA.Element
				set additionalElementToAdd = Repository.GetElementByID(currentElementID) 
				'add element with matching name to the temporary array and remove its name and ID from the name and id array
				temporaryFeatureTypeArray.Add(additionalElementToAdd)
				FeatureTypeNames.RemoveAt(elementNumber)
				FeatureTypeElementIDs.RemoveAt(elementNumber)
			end if
		next
		
		'generate error messages according to content of the temporary array
		dim tempStoredFeatureType AS EA.Element
		if temporaryFeatureTypeArray.count > 1 then
			Session.Output("Error: Found nonunique names for the following classes. [req/uml/feature]")
			'counting one error per name conflict (not one error per class with nonunique name)
			globalErrorCounter = globalErrorCounter + 1
			for each tempStoredFeatureType in temporaryFeatureTypeArray
				dim theFeatureTypePackage AS EA.Package
				set theFeatureTypePackage = Repository.GetPackageByID(tempStoredFeatureType.PackageID) 
				dim theFeatureTypePackageName
				theFeatureTypePackageName = theFeatureTypePackage.Name
				Session.Output("   Class [«"&tempStoredFeatureType.Stereotype&"» "&tempStoredFeatureType.Name&"] in package ["&theFeatureTypePackageName& "]")
			next	
		end if
		
		'get the element with the first elementID and move it to the temporary array
	LOOP
	
 end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Script Name: checkUtkast
' Author: Åsmund Tjora	
' Purpose: check that packages with "Utkast" as part of the package name also has "Utkast" as SOSI_modellstatus tag and that package with the "Utkast"
' SOSI_modellstatus tag also has "Utkast" as part of the name. 
' Date: 10.01.17 
sub checkUtkast(thePackage)
	dim utkastInName, utkastInTag
	'check if "Utkast" is part of the name
	if (len(thePackage.Name)=len(replace(UCase(thePackage.Name),"UTKAST",""))) then utkastInName=false else utkastInName=true
	'check if "utkast" is part of the SOSI_modellstatus tag
	dim taggedValuesCounter
	dim SOSI_modellstatusTag
	dim currentExistingTaggedValue
	SOSI_modellstatusTag = "Missing Tag"
	utkastInTag=false
	for taggedValuesCounter = 0 to thePackage.Element.TaggedValues.Count - 1
		set currentExistingTaggedValue = thePackage.Element.TaggedValues.GetAt(taggedValuesCounter)			
		if currentExistingTaggedValue.Name = "SOSI_modellstatus" then
			if currentExistingTaggedValue.Value = "utkast" then utkastInTag=true
			SOSI_modellstatusTag=currentExistingTaggedValue.Value
		end if
	next
	
	if (utkastInName = true and SOSI_modellstatusTag = "") then
		Session.Output("Error: Package [«"&thePackage.Element.Stereotype&"» "&thePackage.Element.Name& "] has Utkast as part of the name, but the tag [SOSI_modellstatus] has no value. Expected value [utkast]. [/krav/SOSI-modellregister/applikasjonsskjema/standard/pakkenavn/utkast]")
		globalErrorCounter = globalErrorCounter + 1 
	elseif (utkastInName = true and SOSI_modellstatusTag = "Missing Tag") then
		Session.Output("Error: Package [«"&thePackage.Element.Stereotype&"» "&thePackage.Element.Name& "] has Utkast as part of the name, but the tag [SOSI_modellstatus] is missing. [/krav/SOSI-modellregister/applikasjonsskjema/standard/pakkenavn/utkast]")
		globalErrorCounter = globalErrorCounter + 1 	
	elseif (utkastInName=true and utkastInTag=false) then
		Session.Output("Error: Package [«"&thePackage.Element.Stereotype&"» "&thePackage.Element.Name& "] has Utkast as part of the name, but the tag [SOSI_modellstatus] has the value ["&SOSI_modellstatusTag&"]. Expected value [utkast]. [/krav/SOSI-modellregister/applikasjonsskjema/standard/pakkenavn/utkast]")
		globalErrorCounter = globalErrorCounter + 1 
	end if

	if (utkastInName=false and utkastInTag=true) then
		Session.Output("Error: Package [«"&thePackage.Element.Stereotype&"» "&thePackage.Element.Name& "] has [SOSI_modellstatus] tag with utkast value, but Utkast is not part of the package name. [/krav/SOSI-modellregister/applikasjonsskjema/standard/pakkenavn/utkast]")
		globalErrorCounter = globalErrorCounter + 1 
	end if 

	'check case of name.
	if utkastInName and globalLogLevelIsWarning then
		if not(len(replace(thePackage.Name, "Utkast",""))=len(replace(UCase(thePackage.Name),"UTKAST",""))) then
			Session.Output("Warning: Package [«"&thePackage.Element.Stereotype&"» "&thePackage.Element.Name& "]. Unexpected upper/lower case of the Utkast part of the name. [/krav/SOSI-modellregister/applikasjonsskjema/standard/pakkenavn/utkast]")
			globalWarningCounter = globalWarningCounter + 1
		end if
	end if
end sub
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'------------------------------------------------------------START-------------------------------------------------------------------------------------------
' Sub Name: FindInvalidElementsInPackage
' Author: Kent Jonsrud, Magnus Karge...
' Purpose: Main loop iterating all elements in the selected package and conducting tests on those elements

sub FindInvalidElementsInPackage(package) 
			
 	dim elements as EA.Collection 
 	set elements = package.Elements 'collection of elements that belong to this package (classes, notes... BUT NO packages) 
 	Dim myDictionary 
 	dim errorsInFunctionTests 
 			 
 	'check package definition 
 	CheckDefinition(package) 
			 
	'Iso 19103 Requirement 15 - known stereotypes for packages.
	if UCase(package.element.Stereotype) <> "APPLICATIONSCHEMA" and UCase(package.element.Stereotype) <> "LEAF" and UCase(package.element.Stereotype) <> "" then
		if globalLogLevelIsWarning then
			Session.Output("Warning: Unknown package stereotype: [«" &package.element.Stereotype& "» " &package.Name& "]. [/krav/15]")
			globalWarningCounter = globalWarningCounter + 1
		end if	
	end if

	call checkEndingOfPackageName(package)
	call checkUtkast(package)
	
	'Iso 19103 Requirement 16 - unique (NC?)Names on subpackages within the package.
	if ClassAndPackageNames.IndexOf(UCase(package.Name),0) <> -1 then
		Session.Output("Error: Package [" &startPackageName& "] has non-unique subpackage name ["&package.Name&"]. [/krav/16]")				
		globalErrorCounter = globalErrorCounter + 1 
	end if

	ClassAndPackageNames.Add UCase(package.Name)

	'check if the package name is written correctly according to krav/navning
	checkElementName(package)
 			 
	dim packageTaggedValues as EA.Collection 
	set packageTaggedValues = package.Element.TaggedValues 
 			
	'only for applicationSchema packages: 
	'iterate the tagged values collection and check if the applicationSchema package has a tagged value "language" or "designation" with any content [/krav/flerspråklighet/pakke]
	Call checkTVLanguageAndDesignation (package.Element, "language") 
	Call checkTVLanguageAndDesignation (package.Element, "designation")
	'iterate the tagged values collection and check if the applicationSchema package has a tagged value "version" with any content [/req/uml/packaging ]	
	Call checkValueOfTVVersion( package.Element , "version" ) 
	'iterate the tagged values collection and check if the applicationSchema package has a tagged value "SOSI_modellstatus" that is valid [/krav/SOSI-modellregister/ applikasjonsskjema/status]
	Call ValidValueSOSI_modellstatus( package.Element , "SOSI_modellstatus" )
	'iterate the diagrams and checks if there exists one or more diagram names starting with "Hoveddiagram" if not one has been found already [/krav/hoveddiagram/navning]
	if 	not foundHoveddiagram  then
		call CheckPackageForHoveddiagram(package)
	end if 
	'iterate the diagrams in the package and count those named "Hoveddiagram" [/krav/hoveddiagram/detaljering/navning]
	Call FindHoveddiagramsInAS(package)
	call CheckOversiktsdiagram(package)
					
	'check packages' stereotype for right use of lower- and uppercase [/anbefaling/styleGuide] 	
	call checkStereotypes(package)		 
	
	dim packages as EA.Collection 
	set packages = package.Packages 'collection of packages that belong to this package	
			
	'Navigate the package collection and call the FindNonvalidElementsInPackage function for each of them 
	dim p 
	for p = 0 to packages.Count - 1 
		dim currentPackage as EA.Package 
		set currentPackage = packages.GetAt( p ) 
		FindInvalidElementsInPackage(currentPackage) 
				
		'constraints 
		dim constraintPCollection as EA.Collection 
		set constraintPCollection = currentPackage.Element.Constraints 
 			 
		if constraintPCollection.Count > 0 then 
			dim constraintPCounter 
			for constraintPCounter = 0 to constraintPCollection.Count - 1 					 
				dim currentPConstraint as EA.Constraint		 
				set currentPConstraint = constraintPCollection.GetAt(constraintPCounter) 
								
				'check if the package got constraints that lack name or definition (/req/uml/constraint)								
				Call checkConstraint(currentPConstraint, currentPackage)

			next
		end if	
	next 
 			 
 	'------------------------------------------------------------------ 
	'---ELEMENTS--- 
	'------------------------------------------------------------------		 
 			 
	' Navigate the elements collection, pick the classes, find the definitions/notes and do sth. with it 
	'Session.Output( " number of elements in package: " & elements.Count) 
	dim i 
	for i = 0 to elements.Count - 1 
		dim currentElement as EA.Element 
		set currentElement = elements.GetAt( i ) 
				
		'check elements' stereotype for right use of lower- and uppercase [/anbefaling/styleGuide]
		Call checkStereotypes(currentElement)	
 				 
		'Is the currentElement of type Class and stereotype codelist or enumeration, check the initial values are numeric or not (/anbefaling/1)
		if ((currentElement.Type = "Class") and (UCase(currentElement.Stereotype) = "CODELIST"  Or UCase(currentElement.Stereotype) = "ENUMERATION") Or currentElement.Type = "Enumeration") then
			call checkNumericinitialValues(currentElement)
		end if

		' check if inherited stereotypes are all the same
		Call krav14(currentElement)

		' ---ALL CLASSIFIERS---
		'Iso 19103 Requirement 16 - unique NCNames of all properties within the classifier.
		'Inherited properties  also included, strictly not an error situation but implicit redefinition is not well supported anyway
		if currentElement.Type = "Class" or currentElement.Type = "DataType" or currentElement.Type = "Enumeration" or currentElement.Type = "Interface" then
			if ClassAndPackageNames.IndexOf(UCase(currentElement.Name),0) <> -1 then
				Session.Output("Error: Class [«" &currentElement.Stereotype& "» "&currentElement.Name&"] in package: [" &package.Name& "] has non-unique name. [/krav/16]")				
				globalErrorCounter = globalErrorCounter + 1 
			end if

			ClassAndPackageNames.Add UCase(currentElement.Name)

			call krav16unikeNCnavn(currentElement)
		else
			' ---OTHER ARTIFACTS--- Do their names also need to be tested for uniqueness? (need to be different?)
			if currentElement.Type <> "Note" and currentElement.Type <> "Text" and currentElement.Type <> "Boundary" then
				if ClassAndPackageNames.IndexOf(UCase(currentElement.Name),0) <> -1 then
					Session.Output("Debug: Unexpected unknown element with non-unique name [«" &currentElement.Stereotype& "» " &currentElement.Name& "]. EA-type: [" &currentElement.Type& "]. [/krav/16]")
					'This test is dependent on where the artifact is in the test sequence 
				end if
			end if
		end if
				
		'constraints 
		dim constraintCollection as EA.Collection 
		set constraintCollection = currentElement.Constraints 

		if constraintCollection.Count > 0 then 
			dim constraintCounter 
			for constraintCounter = 0 to constraintCollection.Count - 1 					 
				dim currentConstraint as EA.Constraint		 
				set currentConstraint = constraintCollection.GetAt(constraintCounter) 
							
				'check if the constraints lack name or definition (/req/uml/constraint)
				Call checkConstraint(currentConstraint, currentElement)

			next
		end if		



		'If the currentElement is of type Class, Enumeration or DataType continue conducting some tests. If not continue with the next element. 
		if currentElement.Type = "Class" Or currentElement.Type = "Enumeration" Or currentElement.Type = "DataType" then 
 									 
			'------------------------------------------------------------------ 
			'---CLASSES---ENUMERATIONS---DATATYPE  								'   classifiers ???
			'------------------------------------------------------------------		 
			
			'add name and elementID of the featureType (class, datatype, enumeration with stereotype <<featureType>>) to the related array variables in order to check if the names are unique
			if UCase(currentElement.Stereotype) = "FEATURETYPE" then
				FeatureTypeNames.Add(currentElement.Name)
				FeatureTypeElementIDs.Add(currentElement.ElementID)
			end if
			
			'Iso 19103 Requirement 6 - NCNames in codelist codes.
			if (UCase(currentElement.Stereotype) = "CODELIST"  Or UCase(currentElement.Stereotype) = "ENUMERATION" Or currentElement.Type = "Enumeration") then
				call krav6mnemoniskKodenavn(currentElement)
			end if

			'Iso 19103 Requirement 7 - definition of codelist codes.
			if (UCase(currentElement.Stereotype) = "CODELIST"  Or UCase(currentElement.Stereotype) = "ENUMERATION") then
				call krav7kodedefinisjon(currentElement)
			end if
	
			'Iso 19103 Requirement 15 - known stereotypes for classes.
			if UCase(currentElement.Stereotype) = "FEATURETYPE"  Or UCase(currentElement.Stereotype) = "DATATYPE" Or UCase(currentElement.Stereotype) = "UNION" or UCase(currentElement.Stereotype) = "CODELIST"  Or UCase(currentElement.Stereotype) = "ENUMERATION" Or UCase(currentElement.Stereotype) = "ESTIMATED" or UCase(currentElement.Stereotype) = "MESSAGETYPE"  Or UCase(currentElement.Stereotype) = "INTERFACE" then
			else
				if globalLogLevelIsWarning then
					Session.Output("Warning: Class [«" &currentElement.Stereotype& "» " &currentElement.Name& "] has unknown stereotype. [/krav/15]")
					globalWarningCounter = globalWarningCounter + 1
				end if	
			end if

			'Iso 19103 Requirement 15 - known stereotypes for attributes.
			call krav15stereotyper(currentElement)

			'Iso 19109 Requirement /req/uml/profile - well known types. Including Iso 19103 Requirements 22 and 25
			if (UCase(currentElement.Stereotype) = "CODELIST"  Or UCase(currentElement.Stereotype) = "ENUMERATION" Or currentElement.Type = "Enumeration") then
				'codelist code type shall be empty, <none> or <undefined>
			else
				call reqUmlProfile(currentElement)
			end if

			'Iso 19103 Requirement 18 - each classifier must show all its (inherited) properties together in at least one diagram.
			call krav18viseAlt(currentElement)

			'check if there is a definition for the class element (call CheckDefinition function) 
			CheckDefinition(currentElement) 
 										 
			'check if there is there is multiple inheritance for the class element (/krav/enkelArv) 
			'initialize the global variable startClass which is needed in subroutine findMultipleInheritance 
			set startClass = currentElement 
			loopCounterMultipleInheritance = 0
			Call findMultipleInheritance(currentElement) 
 					 
			'check the structure of the value for tag values: designation, description and definition [/krav/flerspråklighet/element]
			if UCase(currentElement.Stereotype) = "FEATURETYPE" then 
				Call structurOfTVforElement( currentElement, "description")
				Call structurOfTVforElement( currentElement, "designation") 
				Call structurOfTVforElement( currentElement, "definition")
			end if 
		
			'check if the class name is written correctly according to krav/navning (name starts with capital letter)
			checkElementName(currentElement)
 											
			if ((currentElement.Type = "Class") and (UCase(currentElement.Stereotype) = "CODELIST")) then
				'Check if an external codelist has a real URL in the codeList tag [/krav/eksternKodeliste]
				Call checkExternalCodelists(currentElement,  "asDictionary")
			end if 
					
					
			dim stereotype
			stereotype = currentElement.Stereotype 
 					
				
			'------------------------------------------------------------------ 
			'---ATTRIBUTES--- 
			'------------------------------------------------------------------					 
 						 
			' Retrieve all attributes for this element 
			dim attributesCollection as EA.Collection 
			set attributesCollection = currentElement.Attributes 
 			 
			if attributesCollection.Count > 0 then 
				dim n 
				for n = 0 to attributesCollection.Count - 1 					 
					dim currentAttribute as EA.Attribute		 
					set currentAttribute = attributesCollection.GetAt(n) 
					'check if the attribute has a definition									 
					'Call the subfunction with currentAttribute as parameter 
					CheckDefinition(currentAttribute) 
					'check the structure of the value for tagged values: designation, description and definition [/krav/flerspråklighet/element]
					Call structurOfTVforElement( currentAttribute, "description")
					Call structurOfTVforElement( currentAttribute, "designation")
					Call structurOfTVforElement( currentAttribute, "definition") 
															
					'check if the attribute's name is written correctly according to krav/navning, meaning attribute name does not start with capital letter
					checkElementName(currentAttribute)
																								
					'constraints 
					dim constraintACollection as EA.Collection 
					set constraintACollection = currentAttribute.Constraints 
 			 
					if constraintACollection.Count > 0 then 
						dim constraintACounter 
						for constraintACounter = 0 to constraintACollection.Count - 1 					 
							dim currentAConstraint as EA.Constraint		 
							set currentAConstraint = constraintACollection.GetAt(constraintACounter) 
									
							'check if the constraints lacks name or definition (/req/uml/constraint)
							Call checkConstraint(currentAConstraint, currentAttribute)

						next
					end if		
				next 
			end if	 
 					 
			'------------------------------------------------------------------ 
			'---ASSOCIATIONS--- 
			'------------------------------------------------------------------ 
 						 
			'retrieve all associations for this element 
			dim connectors as EA.Collection 
			set connectors = currentElement.Connectors 
 					
			'iterate the connectors 
			'Session.Output("Found " & connectors.Count & " connectors for featureType " & currentElement.Name) 
			dim connectorsCounter 
			for connectorsCounter = 0 to connectors.Count - 1 
				dim currentConnector as EA.Connector 
				set currentConnector = connectors.GetAt( connectorsCounter ) 
							
				if currentConnector.Type = "Aggregation" or currentConnector.Type = "Association" then
								
					'target end 
					dim supplierEnd as EA.ConnectorEnd
					set supplierEnd = currentConnector.SupplierEnd
	
					Call structurOfTVforElement(supplierEnd, "description") 
					Call structurOfTVforElement(supplierEnd, "designation")
					Call structurOfTVforElement(supplierEnd, "definition")
									
					'source end 
					dim clientEnd as EA.ConnectorEnd
					set clientEnd = currentConnector.ClientEnd
									
					Call structurOfTVforElement(clientEnd, "description") 
					Call structurOfTVforElement(clientEnd, "designation")
					Call structurOfTVforElement(clientEnd, "definition")
				end if 		
 							
											
				dim sourceElementID 
				sourceElementID = currentConnector.ClientID 
				dim sourceEndNavigable  
				sourceEndNavigable = currentConnector.ClientEnd.Navigable 
				dim sourceEndName 
				sourceEndName = currentConnector.ClientEnd.Role 
				dim sourceEndDefinition 
				sourceEndDefinition = currentConnector.ClientEnd.RoleNote 
				dim sourceEndCardinality		 
				sourceEndCardinality = currentConnector.ClientEnd.Cardinality 
 							 
				dim targetElementID 
				targetElementID = currentConnector.SupplierID 
				dim targetEndNavigable  
				targetEndNavigable = currentConnector.SupplierEnd.Navigable 
				dim targetEndName 
				targetEndName = currentConnector.SupplierEnd.Role 
				dim targetEndDefinition 
				targetEndDefinition = currentConnector.SupplierEnd.RoleNote 
				dim targetEndCardinality 
				targetEndCardinality = currentConnector.SupplierEnd.Cardinality 
 							
				'if the current element is on the connectors client side conduct some tests 
				'(this condition is needed to make sure only associations where the 
				'source end is connected to elements within this applicationSchema package are  
				'checked. Associations with source end connected to elements outside of this 
				'package are possibly locked and not editable) 
				 							 
				dim elementOnOppositeSide as EA.Element 
				if currentElement.ElementID = sourceElementID and not currentConnector.Type = "Realisation" and not currentConnector.Type = "Generalization" then 
					
					'------------------------------------------------------------------ 
					'---'ASSOSIATION'S CONSTRAINTS--- 
					'----START-------------------------------------------------------------- 
					
					dim constraintRCollection as EA.Collection 
					set constraintRCollection = currentConnector.Constraints 
							
					if constraintRCollection.Count > 0 then 
						dim constraintRCounter 
						for constraintRCounter = 0 to constraintRCollection.Count - 1 					 
							dim currentRConstraint as EA.Constraint		 
							set currentRConstraint = constraintRCollection.GetAt(constraintRCounter) 
							'check if the connectors got constraints that lacks name or definition (/req/uml/constraint)
							Call checkConstraint(currentRConstraint, currentConnector)
						next
					end if 
					
					'----END-------------------------------------------------------------- 
					'---'ASSOSIATION'S CONSTRAINTS--- 
					'------------------------------------------------------------------ 
					
					set elementOnOppositeSide = Repository.GetElementByID(targetElementID) 
 								 
					'if the connector has a name (optional according to the rules), check if it starts with capital letter 
					call checkElementName(currentConnector)
					
					'check if elements on both sides of the association are classes with stereotype dataType or of element type DataType
					call krav12(currentElement, currentConnector, elementOnOppositeSide)
													
					'check if there is a definition on navigable ends (navigable association roles) of the connector 
					'Call the subfunction with currentConnector as parameter 
					CheckDefinition(currentConnector) 
 																								 
					'check if there is multiplicity on navigable ends (krav/10)
					call krav10(currentElement, sourceEndNavigable, targetEndNavigable, sourceEndName, targetEndName, sourceEndCardinality, targetEndCardinality)
					 
					'check if there are role names on navigable ends  (krav/11)
					call krav11(currentElement, sourceEndNavigable, targetEndNavigable, sourceEndName, targetEndName, elementOnOppositeSide)
																		 
					'check if role names on connector ends start with lower case (regardless of navigability) (krav/navning)
					call checkRoleNames(currentElement, sourceEndName, targetEndName, elementOnOppositeSide)
					
				end if 
			next 
 						 
			'------------------------------------------------------------------ 
			'---OPERATIONS--- 
			'------------------------------------------------------------------ 
 						 
			' Retrieve all operations for this element 
			dim operationsCollection as EA.Collection 
			set operationsCollection = currentElement.Methods 
 			 
			if operationsCollection.Count > 0 then 
				dim operationCounter 
				for operationCounter = 0 to operationsCollection.Count - 1 					 
					dim currentOperation as EA.Method		 
					set currentOperation = operationsCollection.GetAt(operationCounter) 
 								
					'check the structure of the value for tag values: designation, description and definition [/krav/flerspråklighet/element]
					Call structurOfTVforElement(currentOperation, "description")
					Call structurOfTVforElement(currentOperation, "designation")
					Call structurOfTVforElement(currentOperation, "definition")
								
					'check if the operations's name starts with lower case 
					'TODO: this rule does not apply for constructor operation 
					if not Left(currentOperation.Name,1) = LCase(Left(currentOperation.Name,1)) then 
						Session.Output("Error: Operation name [" & currentOperation.Name & "] in class ["&currentElement.Name&"] shall not start with capital letter. [/krav/navning]") 
						globalErrorCounter = globalErrorCounter + 1 
					end if 
 								 
					'check if there is a definition for the operation (call CheckDefinition function) 
					'call the subroutine with currentOperation as parameter 
					CheckDefinition(currentOperation) 
 																 
				next 
			end if					 
		end if 
  	next 
end sub 
'-------------------------------------------------------------END--------------------------------------------------------------------------------------------


'global variables 
dim globalLogLevelIsWarning 'boolean variable indicating if warning log level has been choosen or not
globalLogLevelIsWarning = true 'default setting for warning log level is true
 
dim startClass as EA.Element  'the class which is the starting point for searching for multiple inheritance in the findMultipleInheritance subroutine 
dim loopCounterMultipleInheritance 'integer value counting number of loops while searching for multiple inheritance
dim foundHoveddiagram 'boolean to check if a diagram named Hoveddiagram is found. If found, foundHoveddiagram = true  
foundHoveddiagram = false 
dim numberOfHoveddiagram 'number of diagrams named Hoveddiagram
numberOfHoveddiagram = 0
dim numberOfHoveddiagramWithAdditionalInformationInTheName 'number of diagrams with a name starting with Hoveddiagram and including additional characters  
numberOfHoveddiagramWithAdditionalInformationInTheName = 0
dim globalErrorCounter 'counter for number of errors 
globalErrorCounter = 0 
dim globalWarningCounter
globalWarningCounter = 0
'Global list of all used names
'http://sparxsystems.com/enterprise_architect_user_guide/12.1/automation_and_scripting/reference.html
dim startPackageName
dim ClassAndPackageNames
Set ClassAndPackageNames = CreateObject("System.Collections.ArrayList")
'Global objects for testing whether a class is showing all its content in at least one diagram. /krav/18
dim startPackage as EA.Package
dim diaoList
dim diagList

'two global variables for checking uniqueness of FeatureType names - shall be updated in sync 
dim FeatureTypeNames 
Set FeatureTypeNames = CreateObject("System.Collections.ArrayList")
dim FeatureTypeElementIDs
Set FeatureTypeElementIDs = CreateObject("System.Collections.ArrayList")

OnProjectBrowserScript 
