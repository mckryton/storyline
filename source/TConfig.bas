Attribute VB_Name = "TConfig"
Option Explicit

Dim m_step_implementations As Collection

Public Property Get StepImplementations() As Collection
    
    Dim step_implementation_class As Variant
    
    If m_step_implementations Is Nothing Then
        Set m_step_implementations = New Collection
        'add all classes with step implementations here:
        For Each step_implementation_class In Array(New steps_unfold_storyline)
            m_step_implementations.Add step_implementation_class
        Next
    End If
    Set StepImplementations = m_step_implementations
End Property

Public Property Let StepImplementations(new_stepimplementations As Collection)
    Set m_step_implementations = new_stepimplementations
End Property

Public Property Get MaxStepFunctionNameLength() As Variant
    
    'VBA compiler crashes under MacOS (v16.41) if a function name is longer than 63 characters
    ' specified max value is 255 characters
    MaxStepFunctionNameLength = 63
End Property
