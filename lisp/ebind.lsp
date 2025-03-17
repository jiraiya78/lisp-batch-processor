(defun C:ebind ()

 (COMMAND "BINDTYPE"  "1")
 (COMMAND "-xref" "b" "*" )
 (COMMAND "AUDIT" "Y")
 (REPEAT 3 (COMMAND "-PURGE" "A" "*" "n"))
 (COMMAND "BINDTYPE"  "0")
 (princ)
)
	
	
(C:ebind)