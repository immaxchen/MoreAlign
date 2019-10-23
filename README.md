# MoreAlign
VBA code for PowerPoint add-in that provides more alignment methods

# Customize Ribbon
customUI/customUI.xml

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
   <ribbon>
     <tabs>
       <tab id="CustomTab" label="My Tab">
         <group id="SampleGroup" label="Sample Group">
           <button id="Button1" label="Stack" imageMso="ChartAreaChart" size="large" onAction="my.xlam!Stack" />
           ...
         </group>
       </tab>
     </tabs>
   </ribbon>
 </customUI>
```

_rels/.rels

```xml
<Relationship Id="someID" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml" />
```

Reference: [Microsoft Doc](https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/customize-the-office-fluent-ribbon-by-using-an-open-xml-formats-file)

# ToDo
stack rows, columns

spread rows, columns

distribute columns horizontal

distribute rows vertical
