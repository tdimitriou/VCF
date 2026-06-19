# VCF Designer - VB.NET Version

This is a VB.NET port of the VCF Designer, a visual drag-and-drop designer for creating user interfaces using the Demac.VCF framework.

## Project Structure

```
.Net/
├── Classes/
│   ├── DesignApplication.vb      - Wrapper for VCF Application (design-time)
│   ├── DesignWindow.vb             - Wrapper for VCF Window (design-time)
│   ├── DesignSurface.vb            - Visual canvas for designing UI
│   ├── DesignerToolbox.vb          - Toolbox with available controls
│   ├── DesignerSelectionManager.vb - Handles selection and hit testing
│   ├── PropertyEditor.vb            - Property grid for selected controls
│   ├── PropertyItem.vb              - Property item data structure
│   └── XAMLWriter.vb               - Serializes VCF objects to XAML
├── Forms/
│   └── frmDesigner.vb              - Main designer form
├── Modules/
│   └── Module1.vb                   - Application entry point
├── VCFDesigner.vbproj               - VB.NET project file
└── README.md                        - This file
```

## Requirements

- .NET Framework 4.8 or later
- Windows Forms
- Demac.VCF.dll (COM interop reference)
- Visual Studio 2019 or later (for building)

## Key Differences from VB6 Version

1. **Windows Forms instead of VB6 PictureBox**: Uses System.Windows.Forms.PictureBox with GDI+ for rendering
2. **Event Handling**: Uses .NET event handlers instead of VB6 event procedures
3. **Graphics**: Uses System.Drawing.Graphics instead of VB6 drawing methods
4. **Collections**: Uses System.Collections.Generic.List instead of VB6 Collection
5. **String Handling**: Uses System.Text.StringBuilder and modern string methods
6. **COM Interop**: Uses .NET COM interop to interact with the VCF framework (VB6 ActiveX DLL)

## Building the Project

1. Ensure Demac.VCF.dll is built and registered
2. Update the reference path in `VCFDesigner.vbproj` if needed
3. Open the project in Visual Studio
4. Build the solution

## Usage

The designer provides:
- **Toolbox**: Drag controls onto the design surface
- **Design Surface**: Visual canvas for arranging controls
- **Property Editor**: Edit properties of selected controls
- **XAML Export**: Save designs as XAML files
- **XAML Import**: Load XAML files into the designer

## Notes

- The VCF framework is a COM-based framework (VB6 ActiveX DLL), so this VB.NET version uses COM interop
- Some VCF types may need explicit casting or COM interop handling
- The designer is designed to work with the existing VCF framework without modifications

