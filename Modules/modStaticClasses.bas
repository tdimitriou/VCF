Attribute VB_Name = "modStaticClasses"
Option Explicit

Public CollectionViewSource             As New CollectionViewSource
Public DependencyPropertiesStatic       As New DependencyPropertiesStatic
Public BindingsManager                  As New BindingsManager
Public Application                      As New ApplicationStatic
Public API                              As New API
Public INIParser                        As New INIParser
Public Mail                             As New Mail
Public Object                           As New ObjectStatic
Public StringConversion                 As New StringConversion
Public NamingManager                    As New NamingManager
Public StringProcessor                  As New StringProcessor
Public Color                            As New VCF.Color
Public Environment                      As New VCF.Environment
Public TypeRegistry                     As New TypeRegistry


' Internal Use Only
Public Conversion                       As New Conversion
