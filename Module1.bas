Attribute VB_Name = "Module1"
Option Explicit
Option Base 0
' --------------------------------- DO NOT SHARE -------------------------------
' --------------------------------- DO NOT SHARE -------------------------------
Public Const accessID As String = "AKIASSOMERANDOMTEXTGOESHERE"
Public Const secretKey As String = "SwirT5EBzcOs6ZwgmWQLITISASECRETOO"
Public Const AcctNum As String = "832175325241"
Public Const queueName As String = "MyQueue"
' --------------------------------- DO NOT SHARE -------------------------------
' --------------------------------- DO NOT SHARE -------------------------------


Sub ReceiveMessages()
Dim WithThis As String, AndThat As String
WithThis = "sqs.us-west-2.amazonaws.com/" & AcctNum & "/" & queueName & "/"
' note 5 second wait at AWS end, program allows 10 seconds for http response at this end
AndThat = "Action=ReceiveMessage&MaxNumberOfMessages=5&VisibilityTimeout=15&AttributeName=All&WaitTimeSeconds=5"
Call AWSAPIGET(WithThis, AndThat)
End Sub

Sub SendMessage()
Dim WithThis As String, AndThat As String
WithThis = "sqs.us-west-2.amazonaws.com/" & AcctNum & "/" & queueName & "/"
' note - program has 10 second wait for response
' also - can't send some characters, like "&" - need to improve parameter splitting!
AndThat = "Action=SendMessage&MessageBody=This is a test message"
Call AWSAPIGET(WithThis, AndThat)
End Sub

Private Function AWSAPIGET(endpoint As String, parameters As String)
' http GET method should not have body content, requires "Host" header
' time now, based on UTC/GMT, used as reference
Dim GMT_Now As Date
GMT_Now = GMT()
Dim dateStamp As String
Dim XAmzDate As String
Dim expires As String
XAmzDate = Format(GMT_Now, "yyyymmdd\Thhnnss\Z")
dateStamp = Left(XAmzDate, 8)
expires = Format(DateAdd("n", 1, GMT_Now), "yyyymmdd\Thhnnss\Z")  ' try 60 seconds after now
' parse endpoint
Dim host As String
Dim service As String
Dim region As String
Dim path As String
host = Split(endpoint, "/")(0)
service = Split(host, ".")(0)
region = Split(host, ".")(1)
path = Mid(endpoint, Len(host) + 1)
' parse query parameters
Dim i As Long
Dim params() As String
Dim paramNames() As String
Dim paramValues() As String
Dim urlParameters As String
' add "Version" and "Expires"
urlParameters = parameters & "&Version=2012-11-05&Expires=" & expires
params = Split(urlParameters, "&")
ReDim paramNames(LBound(params) To UBound(params))
ReDim paramValues(LBound(params) To UBound(params))
For i = LBound(params) To UBound(params)
    paramNames(i) = Split(params(i), "=")(0)
    paramValues(i) = Split(params(i), "=")(1)
Next i
' bubble sort query parameter names to canonical order
' https://wellsr.com/vba/2018/excel/vba-bubble-sort-macro-to-sort-array/
Dim j As Long, tempName As String, tempValue As String
For i = LBound(paramNames) To UBound(paramNames) - 1
    For j = i + 1 To UBound(paramNames)
        If (StrComp(paramNames(i), paramNames(j), vbBinaryCompare) = 1) Then ' swap
            tempName = paramNames(j)
            paramNames(j) = paramNames(i)
            paramNames(i) = tempName
            tempValue = paramValues(j)
            paramValues(j) = paramValues(i)
            paramValues(i) = tempValue
        End If
    Next j
Next i

' follow along with...
' https://docs.aws.amazon.com/general/latest/gr/sigv4-create-canonical-request.html
'
' Task 1, Step 1 - http verb (GET, POST, etc.)
Dim httpRequestMethod As String
httpRequestMethod = "GET"

' Task 1, Step 2: - URI
Dim canonicalUri As String
' double encode EACH SEGMENT of path - use UREyeEncode function, so "/" will not be affected
canonicalUri = UREyeEncode(UREyeEncode(path))

' Task 1, Step 3 - canonical query string
Dim canonicalQueryString As String
canonicalQueryString = ""
For i = LBound(paramNames) To UBound(paramNames)
    canonicalQueryString = canonicalQueryString & QueryUREyeEncode(paramNames(i))
    canonicalQueryString = canonicalQueryString & "=" & QueryUREyeEncode(paramValues(i)) & "&"  ' works with receive message
'  canonicalQueryString = canonicalQueryString & "=" & QueryUREyeEncode(paramValues(i)) & "&amp;"  ' try this
Next i
' strip trailing ampersand
canonicalQueryString = Left(canonicalQueryString, Len(canonicalQueryString) - 1)   ' works with receive message
'canonicalQueryString = Left(canonicalQueryString, Len(canonicalQueryString) - 5)

' Task 1, Step 4 - canonical headers - need to dimension to suit...
' simplify the process - sign ALL the headers (except AUTHORIZATION)
' headers in HTTP rquest are not case sensitive, so just use the lower case version
' For HTTP/1.1 requests, the host header must be included as a signed header
Dim headerNames(0 To 2) As String
Dim headerValues(0 To 2) As String
headerNames(0) = "Host"
headerValues(0) = host
headerNames(1) = "X-Amz-Date"
headerValues(1) = XAmzDate
headerNames(2) = "Content-Type"
headerValues(2) = "application/x-www-form-urlencoded; charset=utf-8"  ' maybe optional, maybe wrong for GET???
' canonize
For i = LBound(headerNames) To UBound(headerNames)
    headerNames(i) = LCase(Trim(headerNames(i)))
    headerValues(i) = TrimAll(headerValues(i))
Next i
' sort
For i = LBound(headerNames) To UBound(headerNames) - 1
    For j = i + 1 To UBound(headerNames)
        If (StrComp(headerNames(i), headerNames(j), vbBinaryCompare) = 1) Then ' swap
            tempName = headerNames(j)
            headerNames(j) = headerNames(i)
            headerNames(i) = tempName
            tempValue = headerValues(j)
            headerValues(j) = headerValues(i)
            headerValues(i) = tempValue
        End If
    Next j
Next i
Dim canonicalHeaders As String
canonicalHeaders = ""
For i = LBound(headerNames) To UBound(headerNames)
    canonicalHeaders = canonicalHeaders & headerNames(i) & ":"
    canonicalHeaders = canonicalHeaders & headerValues(i) & vbLf
Next

' Task 1, Step 5 - signed headers
Dim signedHeaders As String
signedHeaders = ""
For i = LBound(headerNames) To UBound(headerNames)
    signedHeaders = signedHeaders & headerNames(i) & ";"
Next
' strip trailing semi-colon
signedHeaders = Left(signedHeaders, Len(signedHeaders) - 1)

' Task 1, Step 6 - hashed payload
Dim Payload As String, HashedPayload As String
Payload = "" ' that is, the body of the request - we don't have one here
HashedPayload = byte2hex(sha256(Payload))

' Task 1, step 7 - assemble canonical request and hash it
Dim canonicalRequest As String, HashedCanonicalRequest As String
canonicalRequest = httpRequestMethod & vbLf
canonicalRequest = canonicalRequest & canonicalUri & vbLf
canonicalRequest = canonicalRequest & canonicalQueryString & vbLf
canonicalRequest = canonicalRequest & canonicalHeaders & vbLf
canonicalRequest = canonicalRequest & signedHeaders & vbLf
canonicalRequest = canonicalRequest & HashedPayload
HashedCanonicalRequest = byte2hex(sha256(canonicalRequest))

Debug.Print "======================="
Debug.Print "Canonical Request"
Debug.Print "-----------------------"
Debug.Print canonicalRequest
Debug.Print "======================="

' Task 2, string to sign - https://docs.aws.amazon.com/general/latest/gr/sigv4-create-string-to-sign.html
Dim stringToSign As String
stringToSign = "AWS4-HMAC-SHA256" & vbLf
' add request date/time - must match date  in request (x-amz-date)
stringToSign = stringToSign & XAmzDate & vbLf
' add credential scope - date in scope must match date in request (x-amz-date)
Dim credentialScope As String ' we will need this again later
credentialScope = dateStamp & "/" & region & "/" & service & "/aws4_request"
stringToSign = stringToSign + credentialScope & vbLf
' add HashedCanonicalRequest
stringToSign = stringToSign & HashedCanonicalRequest ' no trailing vbLf

Debug.Print "StringToSign"
Debug.Print "-----------------------"
Debug.Print stringToSign
Debug.Print "======================="

' Task 3, Step 1 - signature signing key
Dim SigningKeyString As String, SigningKeyByte() As Byte
Dim FinalRequestSignature As String  ' http request signature
SigningKeyByte = getSignatureKey(secretKey, dateStamp, region, service)
SigningKeyString = byte2hex(SigningKeyByte)
FinalRequestSignature = byte2hex(bHMACSHA256(stringToSign, SigningKeyByte))

' Task 4 - Add the signature to the HTTP request
' supposedly, it can go in the HTTP Header
' generate AuthorizationHeader
Dim AccessKeyID As String
AccessKeyID = accessID
Dim authorizationHeader As String
authorizationHeader = "AWS4-HMAC-SHA256 Credential="
authorizationHeader = authorizationHeader & AccessKeyID & "/" & credentialScope
authorizationHeader = authorizationHeader & ", SignedHeaders=" & signedHeaders
authorizationHeader = authorizationHeader & ", Signature=" & FinalRequestSignature

Debug.Print "Authorization Header"
Debug.Print "-----------------------"
Debug.Print authorizationHeader
Debug.Print "======================="

' url for request
Dim url As String
'url = endpoint & "/" & AcctNum & "/" & queue_name & "/" & query
url = "https://" & endpoint & "?" & urlParameters
Debug.Print "url"
Debug.Print "-----------------------"
Debug.Print url
Debug.Print "======================="

' send the request
'Dim objHTTP As New MSXML2.XMLHTTP60  ' reference: Microsoft XML, v6.0
Dim objHTTP As Object
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP.Open httpRequestMethod, url, False  ' false -> NOT async, waits for reply
For i = LBound(headerNames) To UBound(headerNames)
    objHTTP.setRequestHeader headerNames(i), headerValues(i)
Next i
objHTTP.setRequestHeader "Authorization", authorizationHeader
objHTTP.send
' wait
Application.Wait Now + TimeValue("0:00:10")

Debug.Print vbCrLf & "-----------------"
Debug.Print (objHTTP.Status)

Debug.Print vbCrLf & "-----------------"
Debug.Print (objHTTP.responseText)
Debug.Print "-----------------" & vbCrLf

Set objHTTP = Nothing
End Function

Private Function GMT() As Date
Dim dt As Object
Set dt = CreateObject("WbemScripting.SWbemDateTime")
dt.SetVarDate Now
GMT = dt.GetVarDate(False)
Set dt = Nothing
End Function

Private Function UREyeEncode(StringToEncode As String) As String
' URL encoding, except skips the "/" character (decimal 47)
' adapted from: https://www.freevbcode.com/ShowCode.asp?ID=1512
Dim TempAns As String
Dim CurChr As Integer
CurChr = 1
Do Until CurChr - 1 = Len(StringToEncode)
    Select Case asc(Mid(StringToEncode, CurChr, 1))
        Case 47 To 57, 65 To 90, 97 To 122
            TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
        Case Else
            TempAns = TempAns & "%" & Format(Hex(asc(Mid(StringToEncode, CurChr, 1))), "00")
    End Select
    CurChr = CurChr + 1
Loop
UREyeEncode = TempAns
End Function

Private Function QueryUREyeEncode(StringToEncode As String) As String
' URL encoding, skipping: A-Z, a-z, 0-9, hyphen (-), underscore (_), period (.), and tilde (~)
' does not implement note about double encoding "=" in values
' adapted from: https://www.freevbcode.com/ShowCode.asp?ID=1512
Dim TempAns As String
Dim CurChr As Integer
CurChr = 1
Do Until CurChr - 1 = Len(StringToEncode)
    Select Case asc(Mid(StringToEncode, CurChr, 1))
        Case 45, 46, 48 To 57, 65 To 90, 95, 97 To 122, 126
            TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
        Case Else
            TempAns = TempAns & "%" & Format(Hex(asc(Mid(StringToEncode, CurChr, 1))), "00")
    End Select
    CurChr = CurChr + 1
Loop
QueryUREyeEncode = TempAns
End Function

Private Function TrimAll(ByVal Messy As String)
' trim leading/trailing space, reduce any multiple spaces to a single space
Dim Neat As String
Do
    Neat = Messy
    Messy = Replace(Messy, " ", " ")
Loop Until Neat = Messy
TrimAll = Trim(Neat)
End Function

Private Function getSignatureKey(key As String, dateStamp As String, regionName As String, serviceName As String) As Byte()
Dim kSecret() As Byte, kDate() As Byte, kRegion() As Byte, kService() As Byte, kSigning() As Byte
kSecret = str2byte("AWS4" & key)
kDate = bHMACSHA256(dateStamp, kSecret)
kRegion = bHMACSHA256(regionName, kDate)
kService = bHMACSHA256(serviceName, kRegion)
kSigning = bHMACSHA256("aws4_request", kService)
getSignatureKey = kSigning
End Function

Private Function bHMACSHA256(ByVal sTextToHash As String, ByRef bKey() As Byte) As Byte()
Dim asc As Object, enc As Object
Dim bTextToHash() As Byte
Set asc = CreateObject("System.Text.UTF8Encoding")
Set enc = CreateObject("System.Security.Cryptography.HMACSHA256")
bTextToHash = asc.GetBytes_4(sTextToHash)
enc.key = bKey
bHMACSHA256 = enc.ComputeHash_2((bTextToHash))
Set asc = Nothing
Set enc = Nothing
End Function

Private Function byte2hex(byteArray() As Byte) As String
Dim i As Long
For i = LBound(byteArray) To UBound(byteArray)
    byte2hex = byte2hex & Right(Hex(256 Or byteArray(i)), 2)
Next
byte2hex = LCase(byte2hex)
End Function

Private Function sha256(stringToHash As Variant) As Byte()
Dim ssc As Object
Set ssc = CreateObject("System.Security.Cryptography.SHA256Managed")
sha256 = ssc.ComputeHash_2(str2byte(stringToHash))
Set ssc = Nothing
End Function

Private Function str2byte(s As Variant) As Byte()
If VarType(s) = vbArray + vbByte Then
    str2byte = s
ElseIf VarType(s) = vbString Then
    str2byte = StrConv(s, vbFromUnicode)
Else
    Exit Function
End If
End Function
