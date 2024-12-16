@Code
    Layout = Nothing

    ViewData("Title") = "Titulo Por Defecto"
    ViewData("Message") = "Mensaje por defecto"


End Code





<!DOCTYPE html>

<html>
<head>
    <meta name="viewport" content="width=device-width" />
    <title>@ViewData("Title")</title>
</head>
<body>
    <div>
        <h1>Matrícula</h1>
        <h2>@ViewData("Message")</h2>
    </div>
</body>
</html>
