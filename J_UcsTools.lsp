;|  190904 JanQ
    190904  :   todos   :   listing of saved UCS'
                            inserting Hidden icon-block for each
                            deleting all icon-blocks
|;
(defun C:j_save_ucs ()

    (setq AcDoc (vla-get-ActiveDocument (vlax-get-acad-object)))    ;;---on this document

    (setq Ucs_current
        (vla-add (vla-get-UserCoordinateSystems AcDoc)
        (vlax-3d-point (trans '(0 0 0) 1 0))
        (vlax-3d-point (trans '(1 0 0) 1 0))
        (vlax-3d-point (trans '(0 1 0) 1 0))
        "Ucs_current"                                               ;;---save UCS
)))

(defun C:j_load_ucs ()

    (vla-put-ActiveUcs AcDoc Ucs_current)                           ;;---load saved UCS

)