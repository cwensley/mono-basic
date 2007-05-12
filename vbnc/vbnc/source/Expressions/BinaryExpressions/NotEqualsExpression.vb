' 
' Visual Basic.Net Compiler
' Copyright (C) 2004 - 2007 Rolf Bjarne Kvinge, RKvinge@novell.com
' 
' This library is free software; you can redistribute it and/or
' modify it under the terms of the GNU Lesser General Public
' License as published by the Free Software Foundation; either
' version 2.1 of the License, or (at your option) any later version.
' 
' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Lesser General Public License for more details.
' 
' You should have received a copy of the GNU Lesser General Public
' License along with this library; if not, write to the Free Software
' Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
' 

Public Class NotEqualsExpression
    Inherits BinaryExpression

    Protected Overrides Function GenerateCodeInternal(ByVal Info As EmitInfo) As Boolean
        Dim result As Boolean = True

        ValidateBeforeGenerateCode(Info)

        Dim eqInfo As EmitInfo = Info.Clone(True, False, OperandType)

        result = m_LeftExpression.GenerateCode(eqInfo) AndAlso result
        result = m_RightExpression.GenerateCode(eqInfo) AndAlso result

        Select Case OperandTypeCode
            Case TypeCode.Byte, TypeCode.SByte, TypeCode.Int16, TypeCode.UInt16, TypeCode.Int32, TypeCode.UInt32, TypeCode.Int64, TypeCode.UInt64, TypeCode.Single, TypeCode.Double, TypeCode.Boolean, TypeCode.Char
                Emitter.EmitNotEquals(Info, OperandType)
            Case TypeCode.DateTime
                Emitter.EmitCall(Info, Compiler.TypeCache.System_DateTime__Compare_DateTime_DateTime)
                Emitter.EmitLoadI4Value(Info, 0)
                Emitter.EmitNotEquals(Info, Compiler.TypeCache.System_Int32)
            Case TypeCode.Decimal
                Emitter.EmitCall(Info, Compiler.TypeCache.System_Decimal__Compare_Decimal_Decimal)
                Emitter.EmitLoadI4Value(Info, 0)
                Emitter.EmitnotEquals(Info, Compiler.TypeCache.System_Int32)
            Case TypeCode.Object
                Helper.Assert(Helper.CompareType(OperandType, Compiler.TypeCache.System_Object))
                Emitter.EmitLoadI4Value(Info, Info.IsOptionCompareText)
                Emitter.EmitCall(Info, Compiler.TypeCache.MS_VB_CS_Operators__ConditionalCompareObjectNotEqual_Object_Object_Boolean)
            Case TypeCode.String
                Emitter.EmitLoadI4Value(Info, Info.IsOptionCompareText)
                Emitter.EmitCall(Info, Compiler.TypeCache.MS_VB_CS_Operators__CompareString_String_String_Boolean)
                Emitter.EmitLoadI4Value(Info, 0)
                Emitter.EmitNotEquals(Info, Compiler.TypeCache.System_Int32)
            Case Else
                Throw New InternalException(Me)
        End Select

        Return result
    End Function


    Sub New(ByVal Parent As ParsedObject, ByVal LExp As Expression, ByVal RExp As Expression)
        MyBase.New(Parent, LExp, RExp)
    End Sub

    Public Overrides ReadOnly Property Keyword() As KS
        Get
            Return KS.NotEqual
        End Get
    End Property

    Public Overrides ReadOnly Property IsConstant() As Boolean
        Get
            Return MyBase.IsConstant 'CHECK: is this true?
        End Get
    End Property

    Public Overrides ReadOnly Property ConstantValue() As Object
        Get
            Dim rvalue, lvalue As Object
            lvalue = m_LeftExpression.ConstantValue
            rvalue = m_RightExpression.ConstantValue
            If lvalue Is Nothing Or rvalue Is Nothing Then
                Return Nothing
            Else

                Dim tlvalue, trvalue As Type
                Dim clvalue, crvalue As TypeCode
                tlvalue = lvalue.GetType
                clvalue = Helper.GetTypeCode(Compiler, tlvalue)
                trvalue = rvalue.GetType
                crvalue = Helper.GetTypeCode(Compiler, trvalue)

                If clvalue = TypeCode.Boolean AndAlso crvalue = TypeCode.Boolean Then
                    Return CBool(lvalue) <> CBool(rvalue)
                ElseIf clvalue = TypeCode.DateTime AndAlso crvalue = TypeCode.DateTime Then
                    Return CDate(lvalue) <> CDate(rvalue)
                ElseIf clvalue = TypeCode.Char AndAlso crvalue = TypeCode.Char Then
                    Return CChar(lvalue) <> CChar(rvalue)
                ElseIf clvalue = TypeCode.String AndAlso crvalue = TypeCode.String Then
                    Return CStr(lvalue) <> CStr(rvalue)
                ElseIf clvalue = TypeCode.String AndAlso crvalue = TypeCode.Char OrElse _
                 clvalue = TypeCode.Char AndAlso crvalue = TypeCode.String Then
                    Return CStr(lvalue) <> CStr(rvalue)
                End If

                Dim csmallest As TypeCode
                csmallest = TypeConverter.GetNotEqualsOperandType(clvalue, crvalue)

                Select Case csmallest
                    Case TypeCode.Byte
                        Return CByte(lvalue) <> CByte(rvalue)
                    Case TypeCode.SByte
                        Return CSByte(lvalue) <> CSByte(rvalue)
                    Case TypeCode.Int16
                        Return CShort(lvalue) <> CShort(rvalue)
                    Case TypeCode.UInt16
                        Return CUShort(lvalue) <> CUShort(rvalue)
                    Case TypeCode.Int32
                        Return CInt(lvalue) <> CInt(rvalue)
                    Case TypeCode.UInt32
                        Return CUInt(lvalue) <> CUInt(rvalue)
                    Case TypeCode.Int64
                        Return CLng(lvalue) <> CLng(rvalue)
                    Case TypeCode.UInt64
                        Return CULng(lvalue) <> CULng(rvalue)
                    Case TypeCode.Double
                        Return CDbl(lvalue) <> CDbl(rvalue)
                    Case TypeCode.Single
                        Return CSng(lvalue) <> CSng(rvalue)
                    Case TypeCode.Decimal
                        Return CDec(lvalue) <> CDec(rvalue)
                    Case Else
                        Throw New InternalException(Me)
                End Select
            End If
        End Get
    End Property
End Class
