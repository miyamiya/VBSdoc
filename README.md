VBSdoc
======

Detail
=====
See README.txt

Usage
=====
<pre>
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
</pre>


Example
=====
<pre>
VBSdoc.vbs /i:Text.Util.Class /o:vbsdoc /l:ja /q
</pre>


Complete list of supported tags
=====
<pre>
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
</pre>
