    <#
    .SYNOPSIS
        Send email message.

    .DESCRIPTION
        Send email message. 
        
        Support for:
        - Attachments
        - HTML body
        - Authentication
        - SSL
        
    .PARAMETER SMTPHost
        SMTP server to use for sending the mail.

    .PARAMETER SMTPPort
        SMTP server TCP Port.

    .PARAMETER From
        From mail address.
        
        Use hashtable form to supply a display name.

        This parameter must be one of the following:
        
        [string]    "mailaddress@domain.com"
        [hashtable] @{"mailaddress@domain.com" = "Displayname"}

        If you use a hashtable, it must have exactly 1 item.
        
    .PARAMETER Recipient
        Recipient(s) mail address(es).
        
        This parameter can be either a single item, or an array.

        Each item in the array must be one of the following.
        
        [string]    "mailaddress@domain.com"
        [hashtable] @{"mailaddress@domain.com" = "Displayname"}
        [hashtable] @{
                        "mailaddress1@domain.com" = "Displayname"
                        "mailaddress2@otherdomain.com" = "Displayname"
                    }

        NOTE: for recipient address, hashtables can contain multiple entries.

    .PARAMETER Subject
        Subject of the mail

    .PARAMETER Body
        Body of the mail in text/plain format. Use this parameter if you have formatted
        your mail as text only.
                
    .PARAMETER BodyHTML
        Body of the mail in text/html format. Use this parameter if you have formatted
        your mail as HTML.        

    .PARAMETER AttachmentFile
        One or more files to attach to the mail. This can be either a single string, or
        an array of strings.

    .PARAMETER AttachmentText
        One or more attachments to construct from text. Will be attached as text files.

        This can an array or a single item. Each item must be one of the following:

        [string]    "attachment text"
        [hashtable] @{"filename" = "attachment text"}

        In case of a [string] value, a filename will be constructed using increasing numbers:

        "Attachment_1.txt"
        "Attachment_2.txt"
        ...

    .PARAMETER Username
        Username to use for SMTP server authentication. 
        
        If unspecified, no authentication will be attempted.

    .PARAMETER Password
        Password to use for SMTP server authentication. Can be supplied as [string] or [securestring]
        
    .PARAMETER EnableSSL
        Use SSL when communicating with the SMTP host.

    .EXAMPLE
        Send-Mail -SMTPHost mysmtp.mydomain.com -SMTPPort 25 -From bofh@mydomain.com -Recipient "pooruser@mydomain.com" -Subject "Reboot notification" -Body "The system was rebooted 5 minutes ago fyi"

        Sends normal text mail to pooruser@mydomain.com, using the SMTP mail server at mysmtp.mydomain.com:25.
        Mail addresses are specified using simple strings (no display name included)

    .EXAMPLE
        One of below:
        Send-Mail -SMTPHost mysmtp.mydomain.com -SMTPPort 25 -From @{"bofh@mydomain.com" = "Bastard Operator From Hell"} -Recipient @{"pooruser@mydomain.com" = "Poor user 1"},@{"pooruser2@mydomain.com" = "Poor user 2"} -Subject "Reboot notification" -Body "The system was rebooted 5 minutes ago fyi"

        Send-Mail -SMTPHost mysmtp.mydomain.com -SMTPPort 25 -From @{"bofh@mydomain.com" = "Bastard Operator From Hell"} -Recipient @{"pooruser@mydomain.com" = "Poor user 1"; "pooruser2@mydomain.com" = "Poor user 2"} -Subject "Reboot notification" -Body "The system was rebooted 5 minutes ago fyi"

        Both of the above commands sends normal text mail to pooruser@mydomain.com and pooruser2@mydomain.com, using the SMTP 
        mail server at mysmtp.mydomain.com:25. Mail addresses are specified using hashtables (display name included).

    .EXAMPLE
        $HTML = "<html><body><p>This is a HTML formatted mail - btw the system was rebooted 5 mins ago fyi</p></body></html>"

        Send-Mail -SMTPHost mysmtp.mydomain.com -SMTPPort 25 -From bofh@mydomain.com -Recipient "pooruser@mydomain.com" -Subject "Reboot notification" -BodyHTML $HTML

        Sends HTML formatted mail to pooruser@mydomain.com, using the SMTP mail server at mysmtp.mydomain.com:25. 
        Mail addresses are specified using simple strings (no display name included)

    .EXAMPLE
        $HTML = "<html><body><p>This is a HTML formatted mail - btw the system was rebooted 5 mins ago fyi</p></body></html>"

        Send-Mail -SMTPHost mysmtp.mydomain.com -SMTPPort 25 -From bofh@mydomain.com -Recipient "pooruser@mydomain.com" -Subject "Reboot notification" -BodyHTML $HTML -AttachmentFile "file1.zip","file2.zip"

        Sends HTML formatted mail to pooruser@mydomain.com, using the SMTP mail server at mysmtp.mydomain.com:25. 
        
        Following files are attached to the mail:

        file1.zip
        file2.zip

    .EXAMPLE
        $HTML = "<html><body><p>This is a HTML formatted mail - btw the system was rebooted 5 mins ago fyi</p></body></html>"

        Send-Mail -SMTPHost mysmtp.mydomain.com -SMTPPort 25 -From bofh@mydomain.com -Recipient "pooruser@mydomain.com" -Subject "Reboot notification" -BodyHTML $HTML -AttachmentText "This is a textfile attachment",@{"file.txt" = "This is another text attachment"}

        Sends HTML formatted mail to pooruser@mydomain.com, using the SMTP mail server at mysmtp.mydomain.com:25. 
        
        Following files are attached to the mail:
        
        Attachment_1.txt
        file.txt

        Above is produced from the text in the parameters.

    .OUTPUTS
        Nothing but a status.


    .NOTES
        Author.: Kenneth Nielsen (sharzas @ GitHub.com)
        Version: 1.0

    .LINK
        https://github.com/sharzas/Powershell-Send-Mail
    #>
    
    [CmdletBinding()]

Param (
    [Parameter()]
    [string]$SMTPHost,

    [Parameter()]
    [Int32]$SMTPPort = 25,

    [Parameter()]
    $From,

    [Parameter()]
    $Recipient,

    [Parameter()]
    [string]$Subject = "",

    [Parameter()]
    $Body = "",

    [Parameter()]
    $BodyHTML = "",

    [Parameter()]
    [string[]]$AttachmentFile,

    [Parameter()]
    $AttachmentText,

    [Parameter()]
    [string]$Username = "",

    [Parameter()]
    $Password = $null,

    [Parameter()]
    [switch]$EnableSSL = $false

)






function Resolve-Error
{
       <#
    .SYNOPSIS
        Resolves an Error records various properties, and outputs it in verbose format.

    .DESCRIPTION
        Resolves an Error records various properties, and outputs it in verbose format.
        
        It will also unwind Exception chain.

        NOTE:
        All output will be piped through Out-String, with a preset width, in order to
        make custom logging easier to implement - default width is 160
        
    .PARAMETER ErrorRecord
        The ErrorRecord to resolve, will use last error by default.

    .PARAMETER Width
        The width of the console/output to which the ErrorRecord will be formatted.

        Default is 160

    .EXAMPLE
        Resolve-Error -ErrorRecord $Error[1]

        Will resolve the 2nd error in the list of errors.

    .EXAMPLE
        try {
            Throw "this is a test error!"
        } catch {
            Resolve-Error -ErrorRecord $_
        }

        Will resolve the error that was thrown.

    .EXAMPLE
        try {
            Throw "this is a test error!"
        } catch {
            Resolve-Error -ErrorRecord $_ -Width 200
        }

        Will resolve the error that was thrown width a custom Width

    .OUTPUTS
        ErrorRecord(s) in string formatted output.

    .NOTES
        Author.: Kenneth Nielsen (sharzas @ GitHub.com)
        Version: 1.0

        Credit goes to MSFT Jeffrey Snower who made the source version I used
        as base for this function!

        Link under links to his original version.

    .LINK
        https://devblogs.microsoft.com/powershell/resolve-error/
    #>
    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline)]
        $ErrorRecord=$Error[0],

        [Parameter()]
        $Width = 140
    )

    $ExceptionChainIndent = 3
    
    $Property = ("*",($ErrorRecord|Get-Member -MemberType Properties -Name "HResult"|Where-Object {$_}|ForEach-Object {@{n="HResult"; e={"0x{0:x}" -f $_.HResult}}})|Where-Object {$_})

    $ErrorRecord|Select-Object -Property $Property -ExcludeProperty HResult |Format-List -Force|Out-String -Stream -Width $Width
    $ErrorRecord.InvocationInfo|Format-List *|Out-String -Stream -Width $Width
    $Exception = $ErrorRecord.Exception

    for ($i = 0; $Exception; $i++, ($Exception = $Exception.InnerException))
    {   
        # Build Exception Separator with respect for Width
        $ExceptionSeparator = " [Exception Chain - Exception #{0}] " -f [string]$i
        $ExceptionSeparator = "{0}{1}{2}" -f ("-"*$ExceptionChainIndent),$ExceptionSeparator,("-"*($Width - ($ExceptionChainIndent + $ExceptionSeparator.Length)))

        $ExceptionSeparator|Out-String -Stream -Width $Width
        $Exception|Select-Object -Property *,@{n="HResult"; e={"0x{0:x}" -f $_.HResult}} -ExcludeProperty HResult|Format-List * -Force|Out-String -Stream -Width $Width
    }
} # function Resolve-Error







function Send-Mail
{
    <#
    .SYNOPSIS
        Send email message.

    .DESCRIPTION
        Send email message. 
        
        Support for:
        - Attachments
        - HTML body
        - Authentication
        - SSL
        
    .PARAMETER SMTPHost
        SMTP server to use for sending the mail.

    .PARAMETER SMTPPort
        SMTP server TCP Port.

    .PARAMETER From
        From mail address.
        
        Use hashtable form to supply a display name.

        This parameter must be one of the following:
        
        [string]    "mailaddress@domain.com"
        [hashtable] @{"mailaddress@domain.com" = "Displayname"}

        If you use a hashtable, it must have exactly 1 item.
        
    .PARAMETER Recipient
        Recipient(s) mail address(es).
        
        This parameter can be either a single item, or an array.

        Each item in the array must be one of the following.
        
        [string]    "mailaddress@domain.com"
        [hashtable] @{"mailaddress@domain.com" = "Displayname"}
        [hashtable] @{
                        "mailaddress1@domain.com" = "Displayname"
                        "mailaddress2@otherdomain.com" = "Displayname"
                    }

        NOTE: for recipient address, hashtables can contain multiple entries.

    .PARAMETER Subject
        Subject of the mail

    .PARAMETER Body
        Body of the mail in text/plain format. Use this parameter if you have formatted
        your mail as text only.
                
    .PARAMETER BodyHTML
        Body of the mail in text/html format. Use this parameter if you have formatted
        your mail as HTML.        

    .PARAMETER AttachmentFile
        One or more files to attach to the mail. This can be either a single string, or
        an array of strings.

    .PARAMETER AttachmentText
        One or more attachments to construct from text. Will be attached as text files.

        This can an array or a single item. Each item must be one of the following:

        [string]    "attachment text"
        [hashtable] @{"filename" = "attachment text"}

        In case of a [string] value, a filename will be constructed using increasing numbers:

        "Attachment_1.txt"
        "Attachment_2.txt"
        ...

    .PARAMETER Username
        Username to use for SMTP server authentication. 
        
        If unspecified, no authentication will be attempted.

    .PARAMETER Password
        Password to use for SMTP server authentication. Can be supplied as [string] or [securestring]
        
    .PARAMETER EnableSSL
        Use SSL when communicating with the SMTP host.

    .EXAMPLE
        Send-Mail -SMTPHost mysmtp.mydomain.com -SMTPPort 25 -From bofh@mydomain.com -Recipient "pooruser@mydomain.com" -Subject "Reboot notification" -Body "The system was rebooted 5 minutes ago fyi"

        Sends normal text mail to pooruser@mydomain.com, using the SMTP mail server at mysmtp.mydomain.com:25.
        Mail addresses are specified using simple strings (no display name included)

    .EXAMPLE
        One of below:
        Send-Mail -SMTPHost mysmtp.mydomain.com -SMTPPort 25 -From @{"bofh@mydomain.com" = "Bastard Operator From Hell"} -Recipient @{"pooruser@mydomain.com" = "Poor user 1"},@{"pooruser2@mydomain.com" = "Poor user 2"} -Subject "Reboot notification" -Body "The system was rebooted 5 minutes ago fyi"

        Send-Mail -SMTPHost mysmtp.mydomain.com -SMTPPort 25 -From @{"bofh@mydomain.com" = "Bastard Operator From Hell"} -Recipient @{"pooruser@mydomain.com" = "Poor user 1"; "pooruser2@mydomain.com" = "Poor user 2"} -Subject "Reboot notification" -Body "The system was rebooted 5 minutes ago fyi"

        Both of the above commands sends normal text mail to pooruser@mydomain.com and pooruser2@mydomain.com, using the SMTP 
        mail server at mysmtp.mydomain.com:25. Mail addresses are specified using hashtables (display name included).

    .EXAMPLE
        $HTML = "<html><body><p>This is a HTML formatted mail - btw the system was rebooted 5 mins ago fyi</p></body></html>"

        Send-Mail -SMTPHost mysmtp.mydomain.com -SMTPPort 25 -From bofh@mydomain.com -Recipient "pooruser@mydomain.com" -Subject "Reboot notification" -BodyHTML $HTML

        Sends HTML formatted mail to pooruser@mydomain.com, using the SMTP mail server at mysmtp.mydomain.com:25. 
        Mail addresses are specified using simple strings (no display name included)

    .EXAMPLE
        $HTML = "<html><body><p>This is a HTML formatted mail - btw the system was rebooted 5 mins ago fyi</p></body></html>"

        Send-Mail -SMTPHost mysmtp.mydomain.com -SMTPPort 25 -From bofh@mydomain.com -Recipient "pooruser@mydomain.com" -Subject "Reboot notification" -BodyHTML $HTML -AttachmentFile "file1.zip","file2.zip"

        Sends HTML formatted mail to pooruser@mydomain.com, using the SMTP mail server at mysmtp.mydomain.com:25. 
        
        Following files are attached to the mail:

        file1.zip
        file2.zip

    .EXAMPLE
        $HTML = "<html><body><p>This is a HTML formatted mail - btw the system was rebooted 5 mins ago fyi</p></body></html>"

        Send-Mail -SMTPHost mysmtp.mydomain.com -SMTPPort 25 -From bofh@mydomain.com -Recipient "pooruser@mydomain.com" -Subject "Reboot notification" -BodyHTML $HTML -AttachmentText "This is a textfile attachment",@{"file.txt" = "This is another text attachment"}

        Sends HTML formatted mail to pooruser@mydomain.com, using the SMTP mail server at mysmtp.mydomain.com:25. 
        
        Following files are attached to the mail:
        
        Attachment_1.txt
        file.txt

        Above is produced from the text in the parameters.

    .OUTPUTS
        Nothing but a status.


    .NOTES
        Author.: Kenneth Nielsen (sharzas @ GitHub.com)
        Version: 1.0

    .LINK
        https://github.com/sharzas/Powershell-Send-Mail
    #>
    [CmdletBinding()]

    Param (
        [Parameter(Mandatory = $true)]
        [string]$SMTPHost,

        [Parameter()]
        [Int32]$SMTPPort = 25,

        [Parameter(Mandatory = $true, HelpMessage = 'Format is [HashTable]@{Address = DisplayName} or [String]"Address".' )]
        [ValidateNotNullOrEmpty()]
        $From,

        [Parameter(Mandatory = $true, HelpMessage = 'Format is array of [HashTable]@{Address = DisplayName} or [String]"Address". These can be combined, e.g. @{Address = DisplayName},"Address"' )]
        [ValidateNotNullOrEmpty()]
        $Recipient,

        [Parameter()]
        [string]$Subject = "",

        [Parameter()]
        $Body = "",

        [Parameter()]
        $BodyHTML = "",

        [Parameter()]
        [string[]]$AttachmentFile,

        [Parameter()]
        $AttachmentText,

        [Parameter()]
        [string]$Username = "",

        [Parameter()]
        $Password = $null,

        [Parameter()]
        [switch]$EnableSSL = $false

    )
    Write-Verbose ('Send-Mail(): Invoked')



    function New-ErrorRecord
    {
        <#
        .SYNOPSIS
            Build ErrorRecord from scratch, based on Exception, or based on existing ErrorRecord
    
        .DESCRIPTION
            Build ErrorRecord from scratch, based on Exception, or based on existing ErrorRecord
            
            Especially useful for ErrorRecords used to re-throwing errors in advanced functions
            using $PSCmdlet.ThrowTerminatingError()
    
            Support for:
            - Build ErrorRecord from scratch, exception or existing ErrorRecord.
            - Inheriting InvocationInfo from existing ErrorRecord
            - Adding Exception from existing ErrorRecord, to InnerException chain
              in of Exception in new ErrorRecord, to preserve full Exception history.
            
        .PARAMETER baseObject
            If supplied, the ErrorRecord will be based on this. It must be either of:
    
            [System.Exception]
            ==================
            The ErrorRecord will be created based on the other parameters supplied to the function,
            and this Exception is included as is, as the .Exception property.
    
            [System.Management.Automation.ErrorRecord]
            ==========================================
            The ErrorRecord will be created based on this ErrorRecord. The values of parameters
            that are not supplied to the function, will be derived from this object.
    
            The exception in this object, will be added to .InnerException chain of the Exception
            created for the new ErrorRecord.
    
            If specified, InvocationInfo will be inherited from this object, by storing it in
            the FullyQualifiedErrorId property.
    
        .PARAMETER exceptionType
            If -baseObject has not been supplied, an Exception of this type will be created for
            the ErrorRecord. If this parameter is not supplied, a generic System.Exception will
            be created.
    
        .PARAMETER exceptionMessage
            Message that is added to the new Exception, attached to the ErrorRecord.
            
            
        .PARAMETER errorId
            This is used to construct the FullyQualifiedErrorId.
    
            NOTE:
            If -baseObject is an ErrorRecord, and -InheritInvocationInfo is specified as well,
            this parameter will be overridden with the InvocationInfo.PositionMessage property
            of the existing ErrorRecord.
            
        .PARAMETER errorCategory
            Category set in the CategoryInfo of the ErrorRecord. Must be enumerable via the
            [System.Management.Automation.ErrorCategory] enum.
    
        .PARAMETER targetObject
            Object that was target of the operation. This will be used to display some details
            in the CategoryInfo part of the ErrorRecord - e.g. partial value and Data type of
            the Object (string, int32, etc.).
    
            Hint: it can be a good idea to include this.
    
            NOTE: 
            If -baseObject is an ErrorRecord, and this parameter is not supplied, the 
            targetObject of the existing ErrorRecord will be used in the new ErrorRecord as well,
            unless -DontInheritTargetObject is specified.
                    
        .PARAMETER DontInheritInvocationInfo
            If this parameter is specified, and -baseObject is an ErrorRecord, the InvocationInfo
            will NOT be inherited in the new ErrorRecord.
    
            If this parameter isn't specified, and -baseObject is an ErrorRecord, the 
            InvocationInfo.PositionMessage property of the ErrorRecord in baseObject will be 
            appended to the -errorId parameter supplied to the ErrorRecord constructor.
    
            The benefit from this is, that the resulting ErrorRecord will have correct position
            information displayed in the FullyQualifiedErrorId part of the ErrorRecord, when 
            re-throwing and error in a function, using $PSCmdlet.ThrowTerminatingError()
    
            If this parameter IS supplied, the ErrorRecord will show the position of the 
            exception as the line where the function was called, as opposed to the line where the
            exception was thrown. Not using this parameter includes both positions, so it will be
            possible to see both where the function was called, and where the Exception was thrown.
    
        .PARAMETER DontInheritTargetObject
            If specified, and -baseObject is an ErrorRecord, and -targetObject isn't specified
            either, the value of targetObject will be set to $null, to prevent inheritance of
            this value from the existing ErrorRecord.
    
        .PARAMETER DontUpdateInnerException
            If specified, and -baseObject is an ErrorRecord, the exception created for the new
            ErrorRecord, will not have its .InnerException property chain updated with the
            the Exception from the ErrorRecord in baseObject, and thus the Exception history
            will be reset.
    
        .EXAMPLE
            You have the following advanced functions:
            
            function Test-LevelTwo
            {
                [CmdletBinding()]
                Param ($TestLevelTwoParameter)
    
                try {
                    Get-Content NonExistingFile.txt -ErrorAction Stop
                } catch {
                    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -baseObject $_))
                }
            } # function Test-LevelTwo
    
            function Test-LevelOne
            {
                [CmdletBinding()]
                Param ($TestLevelOneParameter)
    
                try {
                    Test-LevelTwo -TestLevelTwoParameter "This is a parameter for Test-LevelTwo"
                } catch {
                    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -baseObject $_ -DontInheritInvocationInfo))
                }
            } # function Test-LevelOne
    
            And call the function: Test-LevelOne -TestLevelOneParameter "This is a parameter for Test-LevelOne"
    
            It will display the following error:
    
            PS C:\Test> .\Test.ps1
            Test-LevelOne : Cannot find path 'C:\Test\NonExistingFile.txt' because it does not exist.
            At C:\Test\Test.ps1:28 char:1
            + Test-LevelOne -TestLevelOneParameter "This is a parameter for Test-Le ...
            + ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                + CategoryInfo          : ObjectNotFound: (C:\Test\NonExistingFile.txt:String) [Test-LevelOne], ItemNotFoundException
                + FullyQualifiedErrorId : errorId not specified,Test-LevelOne
    
            Note of the error position at line 28, char 1 - this is in fact the line where
            the function Test-LevelOne is called. To get the exact position of the error included 
            in the FullyQualifiedErrorId, do not use the -DontInheritInvocationInfo parameter in
            Test-LevelOne when calling New-ErrorRecord. See next example...
            
        .EXAMPLE
            You have the following advanced function:
            
            function Test-LevelTwo
            {
                [CmdletBinding()]
                Param ($TestLevelTwoParameter)
    
                try {
                    Get-Content NonExistingFile.txt -ErrorAction Stop
                } catch {
                    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -baseObject $_))
                }
            } # function Test-LevelTwo
    
            function Test-LevelOne
            {
                [CmdletBinding()]
                Param ($TestLevelOneParameter)
    
                try {
                    Test-LevelTwo -TestLevelTwoParameter "This is a parameter for Test-LevelTwo"
                } catch {
                    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -baseObject $_))
                }
            } # function Test-LevelOne
    
            And call the function: Test-LevelOne -TestLevelOneParameter "This is a parameter for Test-LevelOne"
    
            It will display the following error:
    
            PS C:\Test> .\Test.ps1
            Test-LevelOne : Cannot find path 'C:\Test\NonExistingFile.txt' because it does not exist.
            At C:\Test\Test.ps1:28 char:1
            + Test-LevelOne -TestLevelOneParameter "This is a parameter for Test-Le ...
            + ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                + CategoryInfo          : ObjectNotFound: (C:\Test\NonExistingFile.txt:String) [Test-LevelOne], ItemNotFoundException
                + FullyQualifiedErrorId : NotSpecified
            +         Source.CategoryInfo     : "ObjectNotFound: (C:\Test\NonExistingFile.txt:String) [Test-LevelTwo], ItemNotFoundException"
            +         Source.Exception.Message: "Cannot find path 'C:\Test\NonExistingFile.txt' because it does not exist."
            +         Source.Exception.Thrown : At C:\Test\Test.ps1:22 char:9
            +         Test-LevelTwo -TestLevelTwoParameter "This is a parameter for ...
            +         ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            +
            --- Test-LevelTwo : NotSpecified
            +         Source.CategoryInfo     : "ObjectNotFound: (C:\Test\NonExistingFile.txt:String) [Get-Content], ItemNotFoundException"
            +         Source.Exception.Message: "Cannot find path 'C:\Test\NonExistingFile.txt' because it does not exist."
            +         Source.Exception.Thrown : At C:\Test\Test.ps1:10 char:9
            +         Get-Content NonExistingFile.txt -ErrorAction Stop
            +         ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            +
            --- Get-Content : PathNotFound,Microsoft.PowerShell.Commands.GetContentCommand,Test-LevelTwo,Test-LevelOne
    
            Take good note of the position for all exceptions in the chain, included
            in FullyQualifiedErrorId - this is the result of -DontInheritInvocationInfo not being used.
    
        .EXAMPLE
            $TestString = "this is a string"
            $ErrorRecord = New-ErrorRecord -exceptionType "System.Exception" -exceptionMessage "This is an Exception" -errorId "This is a test error record" -errorCategory "ReadError" -targetObject $TestString
    
            throw $ErrorRecord
    
            Will throw the following error:
    
            This is an Exception
            At C:\Test\test.ps1:290 char:1
            +     throw $ErrorRecord
            +     ~~~~~~~~~~~~~~~~~~
                + CategoryInfo          : ReadError: (this is a string:String) [], Exception
                + FullyQualifiedErrorId : This is a test error record        
    
        .EXAMPLE
            You may also @Splay parameters via HashTable for better readability:
    
            $TestString = "this is a string value"
    
            $Param = @{
                exceptionType = "System.Exception"
                exceptionMessage = "This is an Exception generated with @Splatted parameters"
                errorId = "This is a test error record"
                errorCategory = "ReadError" 
                targetObject = $TestString
            }
    
            $ErrorRecord = New-ErrorRecord @Param
    
            throw $ErrorRecord
    
            Will throw the following error:
    
            This is an Exception generated with @Splatted parameters
            At C:\Test\test.ps1:290 char:1
            +     throw $ErrorRecord
            +     ~~~~~~~~~~~~~~~~~~
                + CategoryInfo          : ReadError: (this is a string value:String) [], Exception
                + FullyQualifiedErrorId : This is a test error record
    
    
        .OUTPUTS
            An ErrorRecord object.
    
        .NOTES
            Author.: Kenneth Nielsen (sharzas @ GitHub.com)
            Version: 1.2
    
        .LINK
            https://github.com/sharzas/Powershell-New-ErrorRecord
        #>
    
        [CmdletBinding()]
    
        Param (
            [System.Object]$baseObject,
            [System.String]$exceptionType = "System.Exception",
            [System.string]$exceptionMessage = "exceptionMessage not specified",
            [System.string]$errorId = "errorId not specified",
            [System.Management.Automation.ErrorCategory]$errorCategory = "NotSpecified",
            [System.Object]$targetObject = $null,
            [Switch]$DontInheritInvocationInfo,
            [Switch]$DontInheritTargetObject,
            [Switch]$DontUpdateInnerException
        )
    
        Write-Verbose ('New-ErrorRecord(): invoked.')
    
    
    
        function Split-WordWrap
        {
            <#
            .SYNOPSIS
                Word wrap a text block at specified width.
    
            .DESCRIPTION
                Word wrap a text block at specified width.
                
            .PARAMETER Text
                Text block to perform Word Wrap on.
    
            .PARAMETER Width
                Maximum width of each line. Lines will word wrapped so as to not exceed this line
                length.
                
                If any words in the text block is longer than the maximum width of a line, they
                will be split.
    
            .PARAMETER SplitLongWordCharacter
                Character that will be inserted at the end of each line, if it becomes neccessary
                to split a very long word.
                
            .PARAMETER NewLineCharacter
                This is the character that will be used for line feeds. Because of various
                possible scenarious this parameter has been included.
    
                Default is "`n" (LF), but some may prefer "`r`n" (CRLF)
    
                Just be aware that "`r`n" will make Powershell native code interpret an extra
                empty line, for each line... on the other hand, applications that expect "`r`n"
                will need that.
    
            .EXAMPLE
                Split-WordWrap -Line "this is a long line that needs some wrapping" -Width 15
    
                Will output the string value:
    
                this is a long
                line that needs
                some wrapping
                
            .EXAMPLE
                Split-WordWrap -Line "this long line contains SomeVeryLongWordsThatNeedSplitting and SomeOtherVeryLongWords" -Width 15
    
                Will output the string value:
    
                this long line
                contains SomeV-
                eryLongWordsTh-
                atNeedSplitting
                and SomeOtherV-
                eryLongWords
    
            .OUTPUTS
                String containing the input text block, in word wrapped format, according to
                -Width
    
            .NOTES
                Author.: Kenneth Nielsen (sharzas @ GitHub.com)
                Version: 1.0
            #>
    
            [CmdletBinding()]
            Param (
                [String]$Text,
                $Width = $null,
                $SplitLongWordCharacter = "-",
                $NewLineCharacter = "`n"
            )
        
            Write-Verbose ('Split-WordWrap(): invoked')
    
            if ($null -eq $Width) {
                # if Width not supplied, or is null, then simply return as is.
                # should let it work under ISE as well, if calling using some
                # (Get-Host).UI.RawUI values.
                return $Text
            }
    
            # replace single newline characters to CRLF
            $Text = $Text.Replace("`r`n","`n")
        
            # split line into separate lines by CRLF if any is present.
            $Lines = $Text.Split("`n")
        
            $NewContent = foreach ($Line in $Lines) {
                $Words = $Line.Split(" ")
        
                # for each line, start with a blank line variable. We'll add to this one until we reach the specified
                # width, at which point we will wrap to next line.
                $NewLine = ""
        
                foreach ($Word in $Words) {
                    $Skip = $false
            
                    if (($NewLine + ('{0}' -f $Word)).Length -gt $Width) {
                        # Current line + addition of the next word, will exceed the specified width, so we need to wrap here
                        Write-Verbose ('Split-WordWrap(Wrap): ("{0}" + "{1}").Length -gt "{2}" = "{3}"' -f $NewLine, ('{0}' -f $Word), $Width, $(($NewLine + $Word).Length -gt $Width))
    
                        if ($Word.Length -gt $Width) {
                            # The next word is wider than the specified width, so we need to split that word in order to
                            # be able to wrap it.
                            Write-Verbose ('Word is wider than width, need to split in order to wrap: "{0}"' -f $Word)
        
                            $TooLongWord = $Newline + $Word
        
                            Do {
                                $SplittedWord = ('{0}{1}' -f $TooLongWord.Substring(0,($Width-1)), $SplitLongWordCharacter)
                                $SplittedWord
                                Write-Verbose ('Split-WordWrap(): $SplittedWord is now = "{0}"' -f $SplittedWord)
        
                                $TooLongWord = $TooLongWord.Substring($Width-1)
                                Write-Verbose ('Split-WordWrap(): $TooLongWord.Substring({0}) = "{1}"' -f ($Width-1),$TooLongWord)
                            }
                            Until ($TooLongWord.Length -le $Width)
        
                            $NewLine = ('{0} ' -f $TooLongWord)
        
                            # we need to skip adding this word to the current line, as we've just done that.
                            $Skip = $true
                        } else {
                            # The next word is narrower than specified width, so we can wrap simply by completing current
                            # line, and adding this word as the beginning of a new line.
        
                            # output current line
                            Write-Verbose ('Split-WordWrap(): New Line "{0}"' -f $NewLine.Trim())
                            $NewLine.Trim()
        
                            # reset line, in preparation for adding the next word as a new line.
                            $NewLine = ""    
                        }
                    }
        
                    if (!$Skip) {
                        # skip has not been specified, so add current word to current line
                        $NewLine += ('{0} ' -f $Word)
                    }
                    
                }
                Write-Verbose ('Split-WordWrap(): New Line "{0}"' -f $NewLine.Trim())
                $NewLine.Trim()
            }    
        
            Write-Verbose ('Split-WordWrap(): Joining {0} lines to return' -f $NewContent.Count)
    
            $NewContent = $NewContent -Join $NewLineCharacter
            return $NewContent
        } # function Split-WordWrap
    
    
    
        $Record = $null
    
        if ($PSBoundParameters.ContainsKey("baseObject")) {
            # base object was supplied - this must be either [System.Exception] or [System.Management.Automation.ErrorRecord]
            if ($baseObject -is [System.Exception]) {
                # exception
                # an existing exception was specified, so use that to create the errorrecord.
                Write-Verbose ('New-ErrorRecord(): -baseObject is [System.Exception]: build ErrorRecord using this Exception.')
    
                $Record = New-Object System.Management.Automation.ErrorRecord($baseObject, $errorId, $errorCategory, $targetObject)
    
            } elseif ($baseObject -is [System.Management.Automation.ErrorRecord]) {
                # errorrecord
                # an existing ErrorRecord was specified, so use that to create the new errorrecord.
                Write-Verbose ('New-ErrorRecord(): -baseObject is [System.Management.Automation.ErrorRecord]: build ErrorRecord based on this.')
    
                if (!$DontInheritInvocationInfo) {
                    # -DontInheritInvocationInfo NOT specified: construct information about the original invocation, and store it
                    # in errorId of the new record. This is practical if the this errorrecord is made to re-throw via
                    # $PSCmdlet.ThrowTerminatingError in a function. If we don't do this, the ErrorRecord will have invocation
                    # info, and positional info that points to the line in the script, where the function is called from, 
                    # rather than the line where the error occured.
                    Write-Verbose ('New-ErrorRecord(): -DontInheritInvocationInfo NOT specified: Including InvocationInfo.PositionMessage as errorId')
    
                    # Set some indentation values
                    $Indentation = " "*2
                    $DataIndentation = 26
    
                    $PositionMessage = $baseObject.InvocationInfo.PositionMessage.Split("`n") -replace "^\+\s+", ""
    
    
                    if ($PositionMessage.Count -gt 1) {
                        $PositionMessage[1..$PositionMessage.GetUpperBound(0)]|ForEach-Object {
                            $_ = ('+{0}{0}{1}' -f $Indendation,$_)
                        }
                    }
    
                    $PositionMessage = $PositionMessage -join "`n"
    
                    # Base value of errorId
                    $errorIdBase = @'
{0}
    
Source.CategoryInfo     : "{2}"
Source.Exception.Message: "{3}"
Source.Exception.Line   : {4}
Source.Exception.Thrown : {5}
    
--- {6} : {7}
'@
                    
                    if (!$PSBoundParameters.ContainsKey("errorId")) {
                        # -errorId not specified, so simply set value to InvocationInfo.PositionMessage 
                        # of existing ErrorRecord
                        Write-Verbose ('New-ErrorRecord(): -errorId NOT specified: constructing by merging empty string with FullyQualifiedErrorId chain.')
    
                        $errorId = $errorIdBase -f `
                            "NotSpecified", `
                            $Indentation, `
                            (Split-WordWrap -Text $baseObject.CategoryInfo.ToString() -Width ((Get-Host).UI.RawUI.WindowSize.Width - ($DataIndentation+8))).Replace("`n",("`n{0}" -f (" "*$DataIndentation))), `
                            (Split-WordWrap -Text $baseObject.Exception.Message -Width ((Get-Host).UI.RawUI.WindowSize.Width - ($DataIndentation+8))).Replace("`n",("`n{0}" -f (" "*$DataIndentation))), `
                            (Split-WordWrap -Text $baseObject.InvocationInfo.Line.Trim() -Width ((Get-Host).UI.RawUI.WindowSize.Width - ($DataIndentation+8))).Replace("`n",("`n{0}" -f (" "*$DataIndentation))), `
                            (Split-WordWrap -Text $PositionMessage -Width ((Get-Host).UI.RawUI.WindowSize.Width - ($DataIndentation+8))).Replace("`n",("`n{0}+{1}" -f (" "*$DataIndentation), $Indentation)), `
                            $baseObject.InvocationInfo.InvocationName, `
                            $baseObject.FullyQualifiedErrorId
                    } else {
                        # -errorId specified, so merge with existing ErrorRecords InvocationInfo and a NewLine.
                        Write-Verbose ('New-ErrorRecord(): -errorId specified: constructing by merging -errorId with FullyQualifiedErrorId chain.')
    
                        $errorId = $errorIdBase -f `
                            $errorId, `
                            $Indentation, `
                            (Split-WordWrap -Text $baseObject.CategoryInfo.ToString() -Width ((Get-Host).UI.RawUI.WindowSize.Width - ($DataIndentation+8))).Replace("`n",("`n{0}" -f (" "*$DataIndentation))), `
                            (Split-WordWrap -Text $baseObject.Exception.Message -Width ((Get-Host).UI.RawUI.WindowSize.Width - ($DataIndentation+8))).Replace("`n",("`n{0}" -f (" "*$DataIndentation))), `
                            (Split-WordWrap -Text $baseObject.InvocationInfo.Line.Trim() -Width ((Get-Host).UI.RawUI.WindowSize.Width - ($DataIndentation+8))).Replace("`n",("`n{0}" -f (" "*$DataIndentation))), `
                            (Split-WordWrap -Text $PositionMessage -Width ((Get-Host).UI.RawUI.WindowSize.Width - ($DataIndentation+8))).Replace("`n",("`n{0}+{1}" -f (" "*$DataIndentation), $Indentation)), `
                            $baseObject.InvocationInfo.InvocationName, `
                            $baseObject.FullyQualifiedErrorId
                    }
    
                } else {
                    Write-Verbose ('New-ErrorRecord(): -DontInheritInvocationInfo specified: InvocationInfo.PositionMessage not included as errorId')
                }
    
                if (!$PSBoundParameters.ContainsKey("errorCategory")) {
                    # errorCategory wasn't specified, so use the one from the baseObject
                    Write-Verbose ('New-ErrorRecord(): -errorCategory NOT specified: using info from -baseObject ErrorRecord')
    
                    $errorCategory = $baseObject.CategoryInfo.Category
                } else {
                    Write-Verbose ('New-ErrorRecord(): -errorCategory specified: using info from -errorCategory')
                }
    
                Write-Verbose ('New-ErrorRecord(): errorCategory: "{0}"' -f $errorCategory)
    
                if (!$PSBoundParameters.ContainsKey("exceptionMessage")) {
                    # exceptionMessage wasn't specified, so use the one from the exception in the baseObject
                    Write-Verbose ('New-ErrorRecord(): -exceptionMessage NOT specified: using info from -baseObject ErrorRecord')
    
                    $exceptionMessage = $baseObject.exception.message
                }
    
                Write-Verbose ('New-ErrorRecord(): exceptionMessage: "{0}"' -f $errorCategory)
    
                if (!$PSBoundParameters.ContainsKey("targetObject")) {
                    # targetObject wasn't specified
    
                    if ($DontInheritTargetObject) {
                        # -DontInheritTargetObject specified, so set to null
                        Write-Verbose ('New-ErrorRecord(): -targetObject NOT specified, but -DontInheritTargetObject was: setting $null value')
                    } else {
                        # Use the one from the baseObject
                        $targetObject = $baseObject.TargetObject
    
                        Write-Verbose ('New-ErrorRecord(): -targetObject NOT specified: using info from -baseObject ErrorRecord')
                    }
                } else {
                    Write-Verbose ('New-ErrorRecord(): -targetObject specified: added to ErrorRecord')
                }
    
                if ($DontUpdateInnerException) {
                    # Build new exception without adding existing exception from baseObject to InnerException
                    Write-Verbose ('New-ErrorRecord(): -DontUpdateInnerException specified: ErrorRecord Exception will not be added to new Exception.InnerException chain.')
    
                    if ($PSBoundParameters.ContainsKey("exceptionType")) {
                        # -exceptionType specified, use that for the new exception
                        Write-Verbose ('New-ErrorRecord(): -exceptionType specified: creating Exception of type "{0}"' -f $exceptionType)
                        
                        $newException = New-Object $exceptionType($exceptionMessage)
                    } else {
                        # -exceptionType NOT specified, use baseObject.exception type for the new exception
                        Write-Verbose ('New-ErrorRecord(): -exceptionType NOT specified: creating Exception of type "{0}"' -f $baseObject.exception.Gettype().Fullname)
    
                        $newException = New-Object ($baseObject.exception.Gettype().Fullname)($exceptionMessage)
                    }
                } else {
                    # Update InnerException, by adding the exception from the baseObject to the InnerException of the new exception.
                    # this preserves the Exception chain.
                    Write-Verbose ('New-ErrorRecord(): -DontUpdateInnerException NOT specified: ErrorRecord Exception WILL be added to new Exception.InnerException chain.')
    
                    if ($PSBoundParameters.ContainsKey("exceptionType")) {
                        # -exceptionType specified, use that for the new exception
                        Write-Verbose ('New-ErrorRecord(): -exceptionType specified: creating Exception of type "{0}"' -f $exceptionType)
    
                        $newException = New-Object $exceptionType($exceptionMessage, $baseObject.exception)
                    } else {
                        # -exceptionType NOT specified, use baseObject.exception type for the new exception
                        Write-Verbose ('New-ErrorRecord(): -exceptionType NOT specified: creating Exception of type "{0}"' -f $baseObject.exception.Gettype().Fullname)
    
                        $newException = New-Object ($baseObject.exception.Gettype().Fullname)($exceptionMessage, $baseObject.exception)
                    }
                }            
    
                # build the ErrorRecord
        
                Write-Verbose ('New-ErrorRecord(): $newException  = {0}' -f $newException.gettype().fullname)
                Write-Verbose ('New-ErrorRecord(): $errorId       = {0}' -f $errorId.gettype().fullname)
                Write-Verbose ('New-ErrorRecord(): $errorCategory = {0}' -f $errorCategory.gettype().fullname)
                Write-Verbose ('New-ErrorRecord(): $targetObject  = {0}' -f $(if ($null -eq $targetObject) {"null"} else {$targetObject.gettype().fullname}))
    
                $Record = New-Object System.Management.Automation.ErrorRecord($newException, $errorId, $errorCategory, $targetObject)
    
            } else {
                # unsupported type - prepare to create the exception ourselves.
                Write-Verbose ('New-ErrorRecord(): -baseObject is an invalid type [{0}]: will be ignored. Building ErrorRecord using parameters if possible.' -f $baseObject.GetType().FullName)
            }
    
        }
    
        if ($null -eq $Record) {
            # baseObject not specified, or was invalid type, so create ErrorRecord by using parameters
            Write-Verbose ('New-ErrorRecord(): Building ErrorRecord using parameters.')
    
            # output any unspecified parameters verbosely
            @("exceptionMessage","errorId","errorCategory","targetObject")|ForEach-Object {
                if (!$PSBoundParameters.ContainsKey($_)) {
                    # Parameter wasn't specified, use default value.
                    Write-Verbose ('New-ErrorRecord(): -{0} NOT specified: using default value' -f $_)
                }
            }
    
            # create a new exception to embed in the ErrorRecord.
            $newException = New-Object $exceptionType($exceptionMessage)
        
            # Build record
    
            Write-Verbose ('New-ErrorRecord(): $newException  = {0}' -f $newException.gettype().fullname)
            Write-Verbose ('New-ErrorRecord(): $errorId       = {0}' -f $errorId.gettype().fullname)
            Write-Verbose ('New-ErrorRecord(): $errorCategory = {0}' -f $errorCategory.gettype().fullname)
            Write-Verbose ('New-ErrorRecord(): $targetObject  = {0}' -f $(if ($null -eq $targetObject) {"null"} else {$targetObject.gettype().fullname}))
    
            $Record = New-Object System.Management.Automation.ErrorRecord($newException, $errorId, $errorCategory, $targetObject)
        }
    
        # return the constructed ErrorRecord
        $Record
    
    } # function New-ErrorRecord




    $BaseException = @{
        exceptionType = "System.ArgumentException"
        errorId = "Send-Mail.Parameters( -From )"
        errorCategory = "InvalidData"
    }

    # validate from parameter
    if ($From -isnot [Hashtable] -and $From -isnot [string]) {
        $PSCmdlet.ThrowTerminatingError((New-ErrorRecord @BaseException -exceptionMessage '-From is incorrect data type! See help for allowed types!' -targetObject $From))
    }

    if ($From -is [HashTable]) {
        if ($From.Count -ne 1) {
            $PSCmdlet.ThrowTerminatingError((New-ErrorRecord @BaseException -exceptionMessage ('-From = HashTable with {0} entries - count must be exactly 1!' -f $From.Count) -targetObject $From))
        } else {
            Write-Verbose ('Send-Mail(): [Parameter] -From specified as HashTable - {0} items' -f $From.Count)
        }
    } elseif ($From -is [String]) {
        Write-Verbose ('Send-Mail(): [Parameter] -From specified as String')
    }

    # validate recipient parameter
    if ($Recipient -isnot [HashTable] -and $Recipient -isnot [String] -and $Recipient -isnot [String[]] -and $Recipient -isnot [Array]) {
        $BaseException["errorId"] = "Send-Mail.Parameters( -Recipient )"
        $PSCmdlet.ThrowTerminatingError((New-ErrorRecord @BaseException -exceptionMessage '-From is incorrect data type! See help for allowed types!' -targetObject $From))
    }

    $BaseException["errorId"] = "Send-Mail"
    $BaseException.Remove("errorCategory")

    # If the Body parameter is specified, and is an array, join the array to a string value instead
    #
    # This is required in order to use it with the MailMessage class
    #
    If ($Body -ne "" -and $Body -ne $null) {
        If ($Body -is [array]) {
            Write-Verbose ('Send-Mail(): [Parameter] -Body specified as array - converted to String')
            [string]$Body = $Body -join "`n"
        }
    }


    # If the BodyHTML parameter is specified, and is an array, join the array to a string value instead
    #
    # This is required in order to use it with the MailMessage class
    #
    If ($BodyHTML -ne "" -and $BodyHTML -ne $null) {
        If ($BodyHTML -is [array]) {
            Write-Verbose ('Send-Mail(): [Parameter] -BodyHTML specified as array - converted to String')
            [string]$BodyHTML = $BodyHTML -join "`n"
        }
    }



    #
    # Create and set SMTP Client options:
    #
    $SMTPClient = New-Object System.Net.Mail.SmtpClient

    Write-Verbose ('Send-Mail(): [SMTPClient] SMTP Host.........: "{0}"' -f $SMTPHost)
    Write-Verbose ('Send-Mail(): [SMTPClient] SMTP Port.........: "{0}"' -f $SMTPPort)
    Write-Verbose ('Send-Mail(): [SMTPClient] SMTP Enable SSL...: "{0}"' -f $EnableSSL)


    $SMTPClient.Host = $SMTPHost
    $SMTPClient.Port = $SMTPPort
    $SMTPClient.EnableSSL = $EnableSSL


    If ($Username -ne "") {
        #
        # Username specified - so built credentials, and set them on the SMTP client object.
        #
        Write-Verbose ('Send-Mail(): [SMTPClient] Username specified: "{0}"' -f $Username)
        Write-Verbose ('Send-Mail(): [SMTPClient] Password .........: "***"')

        if ($null -ne $Password) {
            if ($Password -isnot [SecureString]) {
                # password specified - but is not SecureString, so convert it.
                $SecPassWd = ConvertTo-SecureString $Password -AsPlainText -Force
            }
        } else {
            # no password specified - so convert empty string to SecureString
            $SecPassWd = ConvertTo-SecureString "" -AsPlainText -Force
        }
        
        $Credentials = New-Object System.Management.Automation.PSCredential ($Username, $SecPassWd)

        $SMTPClient.Credentials = $Credentials.GetNetworkCredential()
    }
    #
    # Done creating SMTP client
    #





    # create new mail message, and set default parameters.
    $Message = New-Object System.Net.Mail.MailMessage

    $Message.BodyTransferEncoding = [System.Net.Mime.TransferEncoding]::Base64
    $Message.BodyEncoding = [System.Text.Encoding]::UTF8
    $Message.HeadersEncoding = [System.Text.Encoding]::UTF8
    $Message.SubjectEncoding = [System.Text.Encoding]::UTF8

    Write-Verbose ('Send-Mail(): [Message] [Subject] ...........: "{0}"' -f $Subject)
    $Message.Subject = $Subject


    if ($From -is [hashtable]) {
        # hashtable - loop through keys (addresses), however ensure only 1 is used.
        foreach ($Address in @($From.Keys)[0]) {
            Write-Verbose ('Send-Mail(): [Message] [Sender]      [HashTable] Adding from: {0} <{1}>' -f $From[$Address], $Address)

            try {
                # set .From / .Sender property 
                $Message.From = (New-Object System.Net.Mail.MailAddress($Address, $From[$Address], [System.Text.Encoding]::UTF8))
                $Message.Sender = (New-Object System.Net.Mail.MailAddress($Address, $From[$Address], [System.Text.Encoding]::UTF8))
            } catch {
                # unable to set sender.
                $PSCmdlet.ThrowTerminatingError((New-ErrorRecord @BaseException `
                    -baseObject $_ `
                    -exceptionMessage ('Error setting .From/.Sender property using the values "{0}", "{1}" - is the format correct? - 0x{2:X} - {3}' -f $Address, $From[$Address], $_.Exception.HResult, $_.Exception.Message) `
                    -errorCategory "InvalidArgument" `
                    -targetObject ('{0} <{1}>' -f $From[$Address], $Address)))
            }
        }

    } elseif ($From -is [string]) {
        # string - use as is.
        Write-Verbose ('Send-Mail(): [Message] [Sender]    [Text]      Adding from: {0} <{1}>' -f $From, $From)

        try {    
            # set .From / .Sender property 
            $Message.From = (New-Object System.Net.Mail.MailAddress($From))
            $Message.Sender = (New-Object System.Net.Mail.MailAddress($From))
        } catch {
            # unable to set sender.
            $PSCmdlet.ThrowTerminatingError((New-ErrorRecord @BaseException `
                -baseObject $_ `
                -exceptionMessage ('Error setting .From/.Sender property using the values "{0}", "{1}" - is the format correct? - 0x{2:X} - {3}' -f $Address, $From[$Address], $_.Exception.HResult, $_.Exception.Message) `
                -errorCategory "InvalidArgument" `
                -targetObject ('{0} <{0}>' -f $From)))
        }
    } else {
        # something else - bail
        Write-Warning ('Send-Mail(): [Message] [Sender]    [ERROR]     Invalid Parameter type specified: {0}' -f $From.Gettype())

        break # bail
    }


    # foreach loop, because there may be more than one Recipient
    #
    foreach ($item in $Recipient) {
        if ($item -is [hashtable]) {
            foreach ($Address in $item.Keys) {
                Write-Verbose ('Send-Mail(): [Message] [Recipient] [HashTable] Adding recipient: {0} <{1}>' -f $item[$Address], $Address)

                try {
                    # attempt to add recipient
                    $Message.To.Add((New-Object System.Net.Mail.MailAddress($Address, $item[$Address], [System.Text.Encoding]::UTF8)))
                } catch {
                    # failed to add recipient
                    $PSCmdlet.ThrowTerminatingError((New-ErrorRecord @BaseException `
                        -baseObject $_ `
                        -exceptionMessage ('Error adding recipient to .To collection using the values "{0}", "{1}" - is the format correct? - 0x{2:X} - {3}' -f $item[$Address], $Address, $_.Exception.HResult, $_.Exception.Message) `
                        -errorCategory "InvalidArgument" `
                        -targetObject ('{0} <{1}>' -f $Address, $item[$Address])))
                }
            }

        } elseif ($item -is [string]) {
            Write-Verbose ('Send-Mail(): [Message] [Recipient] [Text]      Adding recipient: {0} <{1}>' -f $item, $item)

            try {
                # attempt to add recipient
                $Message.To.Add((New-Object System.Net.Mail.MailAddress($item)))
            } catch {
                # failed to add recipient
                Write-Verbose ('Send-Mail(): [Message] [Recipient] [Text]      ERROR: Adding recipient: {0} <{1}>' -f $item, $item)

                $PSCmdlet.ThrowTerminatingError((New-ErrorRecord @BaseException `
                    -baseObject $_ `
                    -exceptionMessage ('Error adding recipient to .To collection using the values "{0}", "{1}" - is the format correct? - 0x{2:X} - {3}' -f $item[$Address], $Address, $_.Exception.HResult, $_.Exception.Message) `
                    -errorCategory "InvalidArgument" `
                    -targetObject ('{0} <{0}>' -f $item)))
            }
        } else {
            Write-Warning ('Send-Mail(): [Message] [Recipient] [ERROR]     Invalid Parameter type specified: {0}' -f $Address.Gettype())
        }
    }



    # If Normal text body is specified, set it.
    #
    If ($Body -ne "" -and $Body -ne $null) {
        Write-Verbose ('Send-Mail(): [Message] Text Body specified.')

        $Message.Body = $Body
    }


    # If HTML body is specified, add it as an alternate view.
    #
    If ($BodyHTML -ne "" -and $BodyHTML -ne $null) {
        Write-Verbose ('Send-Mail(): [Message] HTML Body specified - will add as alternate view')

        # Define mime class to use with the alternate view.
        #
        $mimeHTML = [System.Net.Mime.ContentType][System.Net.Mime.MediaTypeNames+Text]::Html

        $HTMLView = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString($BodyHTML, $mimeHTML)
        
        $Message.AlternateViews.Add($HTMLView)
    }


    If ($PSBoundParameters.Keys.Contains("attachmentfile")) {
        #
        # Some attachment(s) was specified - go grab it/them and attach it to the message
        #
        ForEach ($Item in $AttachmentFile) {
            Write-Verbose ('Send-Mail(): [Attachment] [Filename]  Adding File Attachement - filename: {0}' -f $Item)

            $Att = New-Object System.Net.Mail.Attachment($Item)

            $Message.Attachments.Add($Att)
        }
    }


    If ($PSBoundParameters.Keys.Contains("attachmenttext")) {
        #
        # Some attachment(s) was specified as text - build it/them and attach it to the message
        #
        $i = 0
        ForEach ($Item in $AttachmentText) {

            if ($Item -is [Hashtable]) {
                # this attachment is represented as a HashTable. In that case, it is considered a collection of 1 or more
                # attachments, structed like this:
                #
                # Key   = filename
                # Value = Attachment content in [string] format.
                #
                # @{"filename.txt" = "This is the content of the file."}
                #
                foreach ($Key in $Item.Keys) {
                    Write-Verbose ('Send-Mail(): [Attachment] [HashTable] Adding Text Attachement - filename: {0}' -f $Key)
                    $Att = [System.Net.Mail.Attachment]::CreateAttachmentFromString($Item[$Key], $Key, [System.Text.Encoding]::UTF8, [System.Net.Mime.MediaTypeNames+Text]::Plain)
                }
            } else {
                # this attachment is just a string. In this case a filename is not specified, so we will auto-generate one.
                Write-Verbose ('Send-Mail(): [Attachment] [Text]      Adding Text Attachement - filename: {0}' -f "Attachment_$i.txt")
                $i++
                $Att = [System.Net.Mail.Attachment]::CreateAttachmentFromString($Item, "Attachment_$i.txt", [System.Text.Encoding]::UTF8, [System.Net.Mime.MediaTypeNames+Text]::Plain)
            }

            $Message.Attachments.Add($Att)
        }
    }

    #$Message

    # Send message
    #
    try {
        $SMTPClient.Send($Message)
    } catch {
        $PSCmdlet.ThrowTerminatingError((New-ErrorRecord @BaseException `
            -baseObject $_ `
            -exceptionMessage ('Error calling SMTPClient.Send{0} - 0x{1:X}{0} - {2}' -f "`r`n", $_.Exception.HResult, $_.Exception.Message) `
            -targetObject $Message))
    }
} # function Send-Mail










if ($PSBoundParameters.Count -gt 0) {
    # at least one parameter specified - call Send-Mail interactively    
    Write-Verbose ('Send-Mail.ps1: Parameters specified - adding all non-common ones to Parameter variable.')

    $Params = @{}

    foreach ($key in $PSBoundParameters.Keys) {
        Write-Verbose ('Send-Mail.ps1: [{0}]"{1}" = {2}' -f $PSBoundParameters[$key].Gettype().FullName, $key, $PSBoundParameters[$key])
        $Params[$key] = $PSBoundParameters[$Key]
    }

    @("Verbose","Debug","ErrorAction", "ErrorVariable", "WarningAction", "WarningVariable","OutBuffer", "PipelineVariable", "OutVariable")|Foreach-Object {
        if ($Params.ContainsKey($_)) {
            $Params.Remove($_)
        }
    }
}


if ($Params.Count -gt 0) {
    Write-Verbose ('Send-Mail.ps1: Parameters specified - will call Send-Mail function using below parameters:')
    try {
        Send-Mail @Params

        Write-Host ('Send-Mail.ps1: Mail succesfully sent.')
    } catch {
        Write-Warning ('Send-Mail.ps1: Error sending mail!')

        $_|Resolve-Error|ForEach-Object {Write-Warning "Send-Mail.ps1: $_"}

        # rethrow the statement terminating error, as a script terminating one.
        Throw $_
    }
} else {
    # no parameters specified - do nothing (assume dot sourced / loaded as module)
    Write-Verbose "No non-common parameters specified. Nothing will happen."
    Write-Verbose ""
    Write-Verbose "You may dot source the functions in this script by running it without parameters. If that"
    Write-Verbose "was what you intended, please remember to run the script like this:"
    Write-Verbose ""
    Write-Verbose ". .\Send-Mail.ps1"
}