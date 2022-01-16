<%@ Page Title="" Language="C#" MasterPageFile="~/Website.Master" AutoEventWireup="true" CodeBehind="SchedulePDF.aspx.cs" Inherits="PDF_Demo.View.SchedulePDF" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">

    <div class="legend-width">
        <fieldset>
            <legend><b>Schedule Pdf File</b></legend>
            <table>
                <tr>
                    <td>Select the Pdf File</td>
                    <td>
                        <input type="file" name="postedFile" />
                    </td>
                </tr>
                <tr>
                    <td></td>
                    <td>
                        <input type="button" id="btnUpload" value="Upload" />
                    </td>

                </tr>
            </table>
            <hr />
            <div id="PDFType2" class="tabcontent">
                <div>
                    <h3>Drop Files on Box</h3>
                    <div id="dropOnMe1" draggable="false"></div>
                    <div id="fileCount1" draggable="false"></div>
                    <input id="upload1" draggable="false" type="button" value="Upload Selected Files" />
                    <div draggable="false">
                        <ol draggable="false" id="myFileList1"></ol>
                    </div>
                </div>
            </div>
        </fieldset>
    </div>

    <script>
        /* Code for file upload */
        $(document).ready(function () {
            if (typeof (window.FileReader) == 'undefined') {
                alert('Browser does not support HTML5 file uploads!');
            }

            /* Code for file upload 1 */
            dropOnMe1.addEventListener("drop", dropHandler1, false);

            dropOnMe1.addEventListener("dragover", function (ev) {
                $("#dropOnMe1").css("background-color", "lightgoldenrodyellow;");
                ev.preventDefault();
            }, false);

            function dropHandler1(ev) {
                // Prevent default processing.
                ev.preventDefault();

                // Get the file(s) that are dropped.
                var filelist = ev.dataTransfer.files;
                if (!filelist) return;  // if null, do not do anything.

                $("#dropOnMe1").text(filelist.length +
                    " file(s) selected for uploading!");

                $("#upload1").click(function () {
                    var data = new FormData();
                    for (var i = 0; i < filelist.length; i++) {
                        data.append(filelist[i].name, filelist[i]);
                    }

                    $.ajax({
                        type: "POST",
                        url: "../Services/PdfReaderServiceType2.ashx",
                        contentType: false,
                        processData: false,
                        data: data,
                        success: function (result) {
                            alert(result);
                            location.reload();

                        },
                        error: function () {
                            alert("There was error uploading files!..");
                            location.reload();
                        }
                    });
                });

            }

            dropOnMe1.addEventListener("dragend", function (ev) {
                $("#dropOnMe1").css("background-color", "lightgray;");
                $("#dropOnMe1").text("");
                $("upload1").click(function () { });
                ev.preventDefault();
            }, false);
        });
    </script>
    <script type="text/javascript">
        $("body").on("click", "#btnUpload", function () {
            $.ajax({
                url: "../Services/PdfReaderServiceType2.ashx",
                type: 'POST',
                data: new FormData($('form')[0]),
                cache: false,
                contentType: false,
                processData: false,
                success: function (result) {
                    $("#fileProgress").hide();
                    alert(result);
                    location.reload();
                },
                error: function () {
                    alert("There was error uploading files!..");
                    location.reload();
                }
            });
        });
    </script>
</asp:Content>
