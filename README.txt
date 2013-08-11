Motivation
----------
API documentation is nice, and being able to generate it from the code
is even nicer. However, unlike Perl, Python, Java, or several other
languages, VBScript doesn't have a feature or tool that supports this.
Which kinda sucks.

I tried VBDOX [1], but didn't find usability or results too convincing.
I also tried doxygen [2] by adapting Basti Grembowietz' Visual Basic
doxygen filter. However, doxygen does a lot of things I don't actually
need, and I didn't manage to make it do some of the things I do need.
Thus I ended up writing my own VBScript documentation generator. Enjoy.


Copyright
---------
See COPYING.txt.


Requirements
------------
- VBSdoc uses my Logger class [3] for displaying messages.
- For generation of compiled HTML Help files, HTML Help Workshop [4] must
  be installed and the HTML Help compiler hhc.exe must be in the %PATH%
  of the user running VBSdoc.


Usage
-----
VBSdoc.vbs [/d] [/a] [/q] [/l:LANG] [/p:NAME] [/h:CHM_FILE]
           /i:SOURCE /o:DOC_DIR
VBSdoc.vbs /?

  /?      Print this help.
  /a      Generate documentation for all elements (public and private).
          Without this option, documentation is generated for public
          elements only.
  /d      Enable debug messages. (you really don't want this)
  /h      Create CHM_FILE in addition to normal HTML output. (requires
          HTML Help Workshop)
  /i      Read input files from SOURCE. Can be either a file or a
          directory. (required)
  /l      Generate localized output [de,en]. Default language is en.
  /o      Create output files in DOC_DIR. (required)
  /p      Use NAME as the project name.
  /q      Don't print warnings. Ignored if debug messages are enabled.


Output Format
-------------
The documentation is generated in XHTML format inside the DOC_DIR
directory (see above). For each source file a sub-directory with the
same name (without the extension "vbs") is created, that contains one
or more documentation files. One documentation file for the global
code in the script, and one additional file for each class the script
contains.

Examples: Processing a single script Foo.vbs that contains two classes
          (Bar and Baz) will produce this documentation structure:

          DOC_DIR\
           `- Foo\
               +- Bar.html    <- documentation of class Bar
               +- Baz.html    <- documentation of class Baz
               +- index.html  <- documentation of Foo's global code
               `- vbsdoc.css  <- stylesheet

          Processing a source directory with two scripts Foo.vbs and
          Bar.vbs, of which only Foo.vbs contains a class (Baz) will
          produce these documentation structure:

          DOC_DIR\
           +- Foo\
           |   +- Baz.html    <- documentation of class Baz
           |   `- index.html  <- documentation of Foo's global code
           +- Bar\
           |   `- index.html  <- documentation of Bar's global code
           +- index.html      <- index of global documentation files
           `- vbsdoc.css      <- stylesheet

By default, only code elements with visibility "Public" will be included in
the documentation. Private elements will be omitted, unless the option /a is
used.


Doc Comments
------------
VBSdoc comments begin with the string "'!" (apostrophe followed by an
exclamation mark) and must be placed either before the element they
refer to (without blank lines between doc comment and code) or at the
end of the code line. Examples:

- Valid:
  '! Some procedure.
  '! @param bar Input value
  Sub Foo(bar)

- Valid:
  Const BAR = 42  '! Some constant.
                  '! @see <http://www.example.org/>

- Not valid (blank line between doc comment and code):
  '! Some procedure.
  '! @param bar Input value

    Function Foo

- Not valid (regular comment between doc comment and code):
  '! Some procedure.
  '! @param bar Input value
  ' other comment
  Function Foo

All doc comments for a given code element must be in one consecutive doc
comment block either right before the element, or starting at the end of
the line with the element. Examples:

- Antecedent doc comment:
  '! Some comment.        <- not part of the documentation for Foo()

  '! Some other comment.  <- part of the documentation for Foo()
  Function Foo

- End-of-Line doc comment:
  Const BAR = 42  '! Some End-of-Line   <- part of BAR documentation
                  '! comment.           <- part of BAR documentation

                  '! Some other comment <- not part of BAR documentation

Properties are somewhat special, since they can consist of up to three
functions/procedures (Get/Let/Set). Although it is possible to add doc
comments to each function of a property, I'd recommend to treat all
functions of a property as a single item and add the doc comments to
just one function. Example:

  '! Property Foo of some class.
  '!
  '! @param  index  Index for values of Foo.
  Public Property Get Foo(index)
    Foo = foo_(index)
  End Property

  Public Property Let Foo(index, val)
    foo_(index) = val
  End Property

  '! Property Bar of the same class.
  Public Property Get Bar
    Set Bar = bar_
  End Property

  Public Property Set Bar(obj)
    Set bar_ = obj
  End Property


Tags
----
Tags are used to structure doc comment information. Comment text is
appended to the most recent tag until either the next tag or a blank
comment line appear. A doc comment like this:

  '! @tag This
  '!      is some
  '!      text.
  '!
  '! @other_tag Some other text.

generates the same output as a doc comment like this:

  '! @tag This is some text.
  '! @other_tag Some other text.

Any doc comment that is not associated with a tag is assigned/appended
to the default tag (@details). Detail comments that are separated by
either blank doc comment lines or other tags, become separate paragraphs
in the documentation. Bulleted lists are also supported in detail comments:

  '! Some enumeration:
  '! - item A
  '! - item B
  '! - item C

Example of a properly doc-commented function:

'! Return a slice (sub-array) from a given source array.
'!
'! @param  arr    The source array.
'! @param  first  Index of the beginning of the slice.
'! @param  last   Index of the end of the slice.
'! @return A slice from the given array.
'!
'! @see http://somepage.example.org/
Function Slice(arr, first, last)
  '...
End Function

Complete list of supported tags:

@author   Name and/or mail address of the author. Optional, multiple
          tags per documented element are allowed.

@brief    Summary information. If this tag is omitted, but @details is
          defined, summary information is auto-generated from the first
          sentence or line of the detail information. Should appear at
          most once per documented element.

@date     The release date. Valid for files and classes, otherwise
          ignored. Optional.

@details  Detailed description of the procedure, property or identifyer.
          This is the default tag. The keyword is optional; anything
          that is not associated with any other tag is assigned or
          appended to this tag. If a doc comment does not contain any
          detail description, but does have a summary, the detail
          description is set to the same text as the summary.

@param    Name and description of a function/procedure parameter. Must
          have the form
            @param  NAME  DESCRIPTION
          Where @param-keyword, parameter name and description can be
          separated by any amount of whitespace (except for newlines,
          of course). Valid for functions, procedures, and properties
          with arguments. Multiple tags per documented item are
          allowed.

@raise    Description of the errors a function or procedure might raise.
          Optional, multiple tags per documented element are allowed.
          Valid only for procedures/functions (including properties).

@return   Description of the return value of a function. Required for
          functions, must not appear with any other element. Must not
          appear more than once.

@see      Link to some other resource (external or internal). External
          references should be given as URLs (e.g. http://example.org/)
          and may be enclosed in angular brackets. Descriptive text may
          be placed after the reference:
            @see  REF  DESCRIPTION
          Optional, multiple tags per documented element are allowed.

@todo     An unfinished task. @todo doc comments are somewhat special,
          as they are extracted from source files before the processing
          of the actual code elements. They're grouped into one list
          per source file that is placed at the beginning of the main
          documentation file for that source file. Optional.

@version  Version number. Valid for files and classes, otherwise
          ignored. Optional.

Sensible documentation for any given element should have at least a summary
or a detail description. If both are missing, a warning will be issued,
although the documentation generation will continue.


References
----------
[1] http://vbdox.sourceforge.net/
[2] http://www.doxygen.org/
[3] http://www.planetcobalt.net/download/LoggerClass-1.2.zip
[4] http://msdn.microsoft.com/en-us/library/ms670169.aspx
