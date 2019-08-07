
<!DOCTYPE html>
<html>
    <head>
        <title>Docx2Html By Habib Endris</title>
        <script src="jquery-3.4.1.js"></script>
    </head>
    <body>
    <?php
        require("docx.class.php");

        $docx = new Doc2Txt("test.docx");
        $docx->convertToText();

    // $a1=array("a"=>"red","b"=>"green","c"=>"blue","d"=>"yellow");
    // $a2=array();
    // print_r(array_splice($a1,1,1));
    // echo "<br>";
    // var_dump($a1);
    ?>
    </body>
</html>