# UserForms-Class-Spliter VBA Project

## Description

The project is a library for creating splitters (dividers) between controls on a UserForm in Excel. Splitters allow users to dynamically change the size of controls on a form.

## Project Structure

- `clsSpliter.cls` - main splitter manager class
- `clsSpliterItem.cls` - class for a single splitter item
- `frmSpliter.frm` - form with splitter usage example
- `modShowForms.bas` - module for form display

## Main Components

### clsSpliter
- Manages a collection of splitters
- Provides methods for adding, removing, and changing splitter properties
- Supports bulk operations on all splitters

### clsSpliterItem
- Represents a single splitter item
- Manages two controls that it divides
- Handles splitter movement and element size changes

### frmSpliter
- Form with splitter usage example
- Demonstrates library functionality
- Contains controls for splitter testing

## Usage

### Creating a Splitter
```vba
Dim spliter As New clsSpliter
Call spliter.Initialize(UserForm1) ' or another parent object
Call spliter.AddItem(Control1, Control2) ' create splitter between two controls
```

### Managing Splitters
```vba
' Set text color for all splitters
Call spliter.setForeColorAll(vbBlue)

' Set caption for all splitters
Call spliter.setCaptionAll("Divider")

' Manage visibility
Call spliter.setVisibleAll(True)

' Manage availability
Call spliter.setEnabledAll(True)

' Remove all splitters
Call spliter.RemoveAll()
```

### Managing Individual Splitter
```vba
' Get splitter by index
Dim item As clsSpliterItem
Set item = spliter.item(1)

' Change individual splitter properties
item.Caption = "New Caption"
item.ForeColor = vbRed
item.Visible = True
item.Enabled = True

' Or using clsSpliter methods
Call spliter.ByItemSetCaption(1, "Caption")
Call spliter.ByItemSetForeColor(1, vbRed)
Call spliter.ByItemSetVisible(1, True)
Call spliter.ByItemSetEnabled(1, True)
```

## Testing

To test the project functionality, you can use the frmSpliter form:
1. Run the `showForm` macro from the `modShowForms` module
2. Use form controls to check splitter functionality

## Compatibility

The project is designed for use with Microsoft Excel and VBA. Compatible with Excel 2007 and later versions.

## License

Apache License