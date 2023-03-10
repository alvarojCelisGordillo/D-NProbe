#pragma checksum "C:\Users\Nelly\source\repos\DNPrueba\DNPrueba.Web\Views\Report\Index.cshtml" "{ff1816ec-aa5e-4d10-87f7-6f4963833460}" "7a75490278cebf0f584185320d6c57f9ecd03001"
// <auto-generated/>
#pragma warning disable 1591
[assembly: global::Microsoft.AspNetCore.Razor.Hosting.RazorCompiledItemAttribute(typeof(AspNetCore.Views_Report_Index), @"mvc.1.0.view", @"/Views/Report/Index.cshtml")]
namespace AspNetCore
{
    #line hidden
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.AspNetCore.Mvc.Rendering;
    using Microsoft.AspNetCore.Mvc.ViewFeatures;
#nullable restore
#line 1 "C:\Users\Nelly\source\repos\DNPrueba\DNPrueba.Web\Views\_ViewImports.cshtml"
using DNPrueba.Web;

#line default
#line hidden
#nullable disable
#nullable restore
#line 2 "C:\Users\Nelly\source\repos\DNPrueba\DNPrueba.Web\Views\_ViewImports.cshtml"
using DNPrueba.Web.Models;

#line default
#line hidden
#nullable disable
    [global::Microsoft.AspNetCore.Razor.Hosting.RazorSourceChecksumAttribute(@"SHA1", @"7a75490278cebf0f584185320d6c57f9ecd03001", @"/Views/Report/Index.cshtml")]
    [global::Microsoft.AspNetCore.Razor.Hosting.RazorSourceChecksumAttribute(@"SHA1", @"7bf7d71c1ab00bc3fc4300bc7905daf255952b43", @"/Views/_ViewImports.cshtml")]
    public class Views_Report_Index : global::Microsoft.AspNetCore.Mvc.Razor.RazorPage<dynamic>
    {
        #pragma warning disable 1998
        public async override global::System.Threading.Tasks.Task ExecuteAsync()
        {
            WriteLiteral(@"
<style>

    .content {
        margin-top: 98px;
        text-align: center;
    }
    .content .heading__div {
        text-align: center;
    }

    .content .btnDwn {
        width: 240px;
        height: 52px;
        outline: none;
        border: none;
        background-color: blue;
        color: #fff;
        font-size: 16px;
        font-weight: 600;
        margin-top: 50px;
    }
</style>

<div class=""content"">
    <div class=""heading__div"">
        <h5>Ahora puedes obtener un reporte anual de manera r??pida y eficaz. Has click en el bot??n inferior para descargarlo ahora. </h5>
    </div>
    <button class=""btnDwn"" asp->Descargar Reporte</button>
</div>


");
            DefineSection("Scripts", async() => {
                WriteLiteral(@"
    <script>
        $(document).ready(function() {
            $('.btnDwn').on('click',
                function() {
                    $.ajax({
                        cache: false,
                        type: ""POST"",
                        url: '");
#nullable restore
#line 42 "C:\Users\Nelly\source\repos\DNPrueba\DNPrueba.Web\Views\Report\Index.cshtml"
                         Write(Url.Action("MakeReport"));

#line default
#line hidden
#nullable disable
                WriteLiteral(@"',
                        success: function(data) {
                            var linkSource = `data:application/vnd.ms-excel;base64,${data}`;
                            var downloadLink = document.createElement('a');
                            var filename = ""Reporte.xlsx"";

                            downloadLink.href = linkSource;
                            downloadLink.download = filename;
                            downloadLink.click();
                        },
                        error: function(err) {
                            console.log(err);
                        }
                    });
                });


            function Base64ToBytes(base64) {
                var s = window.atob(base64);
                var bytes = new Uint8Array(s.length);
                for (var i = 0; i < s.length; i++) {
                    bytes[i] = s.charCodeAt(i);
                }
                return bytes;
            };
        });
    </script>
");
            }
            );
        }
        #pragma warning restore 1998
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.ViewFeatures.IModelExpressionProvider ModelExpressionProvider { get; private set; }
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.IUrlHelper Url { get; private set; }
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.IViewComponentHelper Component { get; private set; }
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.Rendering.IJsonHelper Json { get; private set; }
        [global::Microsoft.AspNetCore.Mvc.Razor.Internal.RazorInjectAttribute]
        public global::Microsoft.AspNetCore.Mvc.Rendering.IHtmlHelper<dynamic> Html { get; private set; }
    }
}
#pragma warning restore 1591
