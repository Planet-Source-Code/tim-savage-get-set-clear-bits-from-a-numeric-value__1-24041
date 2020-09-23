Attribute VB_Name = "mdlBits"
'**********************************************************************
'* Description: Modual that manipulates bits in an integer value,
'*              please note that bits are numbered from 0.
'*		Note also that you can change the datatype to any other
'*		non-floating point numberic datatype (ie long, byte).
'*
'* Project:     n/a
'*
'* File:        mdlBits.bas
'*
'* Revision History:
'*    13/06/2001  Tim Savage  Initial Version
'*
'**********************************************************************
Option Explicit

'****************************************************************************************
'* PUBLIC FUNCTIONS
'****************************************************************************************
'******************************************************************************
'* Function:    GetBit
'* Description: Gets a specific bit out of an integer
'* Parameters:  bytBit: Specific bit to return
'*              iValue: Integer value to pull bit out of.
'* Returns:     Boolean indicating the state of the returned bit (ie True = 1,
'*              False = 0)
'******************************************************************************
Public Function GetBit(bytBit As Byte, iValue As Integer) As Boolean
  GetBit = (iValue And 2 ^ bytBit) > 0
End Function

'******************************************************************************
'* Function:    SetBit
'* Description: Well take a integer and set a specific bit
'* Parameters:  bytBit: Specific bit to set
'*              iValue: Integer value to set bit in.
'* Returns:     An Integer containing the value of the input with the specific
'*              bit set (ie if the functions is called SetBit(2, 24), 28 is
'*              returned. 24 is equivilent to 11000 and 28 is 11010.
'******************************************************************************
Public Function SetBit(bytBit As Byte, iValue As Integer) As Integer
  SetBit = iValue Or 2 ^ bytBit
End Function

'******************************************************************************
'* Function:    ClearBit
'* Description: Is similar to Set bit but instead sets a bit to 0
'* Parameters:  bytBit: Specific bit to clear
'*              iValue: Integer value to clear bit in.
'* Returns:     An Integer containing the value of the input with the specific
'*              bit cleared (ie if the functions is called SetBit(2, 28), 24 is
'*              returned.
'******************************************************************************
Public Function ClearBit(bytBit As Byte, iValue As Integer) As Integer
  ClearBit = iValue Xor 2 ^ bytBit
End Function
