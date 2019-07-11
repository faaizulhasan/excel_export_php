<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport"
          content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Document</title>
</head>
<body>
<h2>Import Excel File into MySQL Database using PHP</h2>

<div class="outer-container">
    <form action="" method="post"
          name="frmExcelImport" id="frmExcelImport" enctype="multipart/form-data">
        <div>
            <label>Choose Excel
                File</label> <input type="file" name="file"
                                    id="file" accept=".xls,.xlsx">
            <button type="submit" id="submit" name="import"
                    class="btn-submit">Import</button>

        </div>

    </form>

</div>







<?php
$conn = mysqli_connect("localhost","root","","test");
require_once('php-excel-reader/excel_reader2.php');
require_once('SpreadsheetReader.php');

if (isset($_POST["import"]))
{

    $allowedFileType = ['application/vnd.ms-excel','text/xls','text/xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];

    if(in_array($_FILES["file"]["type"],$allowedFileType)){

        $targetPath = 'uploads/'.$_FILES['file']['name'];
        move_uploaded_file($_FILES['file']['tmp_name'], $targetPath);

        $Reader = new SpreadsheetReader($targetPath);

        $sheetCount = count($Reader->sheets());

        for($i=0;$i<$sheetCount;$i++)
        {
            $Reader->ChangeSheet($i);

            foreach ($Reader as $Row)
            {

                $user_name = "";
                if(isset($Row[0])) {
                    $user_name = mysqli_real_escape_string($conn,$Row[0]);
                    $email = strtolower(str_replace(' ', '', $user_name)).'@gmail.com';
                }

                $denomination = "";
                if(isset($Row[1])) {
                    $denomination = mysqli_real_escape_string($conn,$Row[1]);
                }

                $address = "";
                if(isset($Row[2])) {
                    $address = mysqli_real_escape_string($conn,$Row[2]);
                }

                if (!empty($user_name) || !empty($denomination)|| !empty($email) || !empty($address)) {
                    $query = "insert into users(UserName,Email,Password,Address,ChurchDenomination,UserType) values('".$user_name."','".$email."','".md5('123456')."','".$address."','".$denomination."','".'Church'."')";
                    $result = mysqli_query($conn, $query);

                    if (! empty($result)) {
                        $type = "success";
                        $message = "Excel Data Imported into the Database";
                    } else {
                        $type = "error";
                        $message = "Problem in Importing Excel Data";
                    }
                }
            }

        }
    }
    else
    {
        $type = "error";
        $message = "Invalid File Type. Upload Excel File.";
    }
}
?>







</body>
</html>