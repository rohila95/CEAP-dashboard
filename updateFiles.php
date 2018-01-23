<?php
// echo $_REQUEST['files'];
// echo "hello";

$uploaddir = '/var/www/qav2/kumar/OduEfp/';
$uploadfile = $uploaddir . basename($_FILES['userFile']['name']);
echo $uploadfile;
echo '<pre>';
if (move_uploaded_file($_FILES['userFile']['tmp_name'], $uploadfile)) {
    echo "File is valid, and was successfully uploaded.\n";
} else {
    echo "Possible file upload attack!\n";
}

echo 'Here is some more debugging info:';
print_r($_FILES);

print "</pre>";
?>