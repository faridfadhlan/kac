const app = new Vue({
    el: '#app',
    data: {
        file: null,
        inputKaizala: [],
        absens: [],
        timezone: 7,
        headerText: {
            "responderName": {
                id: 0,
                text: "Responder Name"
            },
            "groupName": {
                id: 1,
                text: "Group Name"
            },
            "responderLocationLatitude": {
                id: 2,
                text: "Responder Location Latitude"
            },
            "responderLocationLongitude": {
                id: 3,
                text: "Responder Location Longitude"
            },
            "responderLocationLocation": {
                id: 4,
                text: "Responder Location Location"
            },
            "responseTime": {
                id: 5,
                text: "Response Time (UTC)"
            },
            "notesQuestionTitle": {
                id: 6,
                text: "NotesQuestionTitle"
            },
            "serverReceiptTimestamp": {
                id: 7,
                text: "Server Receipt Timestamp (UTC)"
            },
        }
    },
    methods: {
        handleFile: (e) => {
            e.preventDefault();
            app.file = e.target.files[0];
        },
        convert: () => {
            Papa.parse(app.file, {
                header: true,
                complete: (results, f) => {
                    app.inputKaizala = results.data;
                    app.addUnixTime();
                    app.sortAbsen();
                    app.processConvert();
                    // app.writeExcel()
                }
            });
        },
        addUnixTime: () => {
            app.inputKaizala = app.inputKaizala.filter(a => {
                return a["Responder Name"] != "";
            })
            app.inputKaizala = app.inputKaizala.map(d => {
                d.unixservertime = moment(d['Server Receipt Timestamp (UTC)']).add(app.timezone, 'h').unix();
                return d;
            })
        },
        sortAbsen: () => {
            app.inputKaizala = _.sortBy(app.inputKaizala, ['unixservertime']);
        },
        processConvert: () => {
            const unik = _.groupBy(app.inputKaizala, 'Responder Name');
            const absen_per_orang = Object.keys(unik).map(k => {
                const absens = {
                    nama: k,
                    datang: unik[k][0],
                    pulang: unik[k].length > 1 ? unik[k][1] : false
                }
                return absens;
            });

            
            app.absens = absen_per_orang.map(apo => {
                const data = {};
                const pecahnama = apo.nama.split('-');
                data.nama = pecahnama[0].trim();
                data.nip = pecahnama[1].trim();
                data.datang = {
                    waktu: apo.datang.unixservertime,
                    latitude: apo.datang['Responder Location Latitude'],
                    longitude: apo.datang['Responder Location Longitude'],
                    location: apo.datang['Responder Location Location'],
                    latlong: apo.datang['Responder Location Latitude'] + ', ' + apo.datang['Responder Location Longitude']
                };
                if(apo.pulang) {
                    data.pulang = {
                        waktu: apo.pulang.unixservertime,
                        latitude: apo.pulang['Responder Location Latitude'],
                        longitude: apo.pulang['Responder Location Longitude'],
                        latlong: apo.pulang['Responder Location Latitude'] + ', ' + apo.pulang['Responder Location Longitude'],
                        location: apo.pulang['Responder Location Location']
                    }
                }
                else {
                    data.pulang = false
                }
                
                
                return data;
            })
            app.writeExcel();

            // $('#tabel-absen').modal('show')
            
        },
        writeExcel: () => {
            const workbook = new ExcelJS.Workbook();
            workbook.creator = 'BPS';
            workbook.lastModifiedBy = 'BPS';

            const worksheet = workbook.addWorksheet('Tabel Absensi');

            worksheet.getColumn(1).width = 15;
            worksheet.getColumn(2).width = 30;
            worksheet.getColumn(3).width = 20;
            worksheet.getColumn(4).width = 25;
            worksheet.getColumn(5).width = 40;
            worksheet.getColumn(6).width = 20;
            worksheet.getColumn(7).width = 25;
            worksheet.getColumn(8).width = 40;

            // kolom NIP
            worksheet.mergeCells('A1:A2');
            worksheet.getCell('A1').value = 'NIP';
            worksheet.getCell('A1').width = 15;
            worksheet.getCell('A1').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('A1').font = { bold: true };

            // kolom nama
            worksheet.mergeCells('B1:B2');
            worksheet.getCell('B1').value = 'Nama';
            worksheet.getCell('B1').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('B1').font = { bold: true };

            // kolom absen datang
            worksheet.mergeCells('C1:E1');
            worksheet.getCell('C1').value = 'Absen Datang';
            worksheet.getCell('C1').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('C1').font = { bold: true };

            // kolom absen pulang
            worksheet.mergeCells('F1:H1');
            worksheet.getCell('G1').value = 'Absen Pulang'
            worksheet.getCell('G1').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('G1').font = { bold: true };

            worksheet.getCell('C2').value = 'Waktu';
            worksheet.getCell('C2').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('C2').font = { bold: true };
            worksheet.getCell('D2').value = 'Latlong';
            worksheet.getCell('D2').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('D2').font = { bold: true };
            worksheet.getCell('E2').value = 'Alamat';
            worksheet.getCell('E2').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('E2').font = { bold: true };

            worksheet.getCell('F2').value = 'Waktu';
            worksheet.getCell('F2').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('F2').font = { bold: true };
            worksheet.getCell('G2').value = 'Latlong';
            worksheet.getCell('G2').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('G2').font = { bold: true };
            worksheet.getCell('H2').value = 'Alamat';
            worksheet.getCell('H2').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('H2').font = { bold: true };
            

            worksheet.getRow(2).commit(); // now rows 1 and two are committed.
            // console.log(app.absens);return;

            for(i in app.absens) {
                const item = app.absens[i];
                if(item.pulang) {
                    worksheet.addRow([
                        item.nip,
                        item.nama,
                        new Date(item.datang.waktu * 1000),
                        item.datang.latlong,
                        item.datang.location,
                        new Date(item.pulang.waktu * 1000),
                        item.pulang.latlong,
                        item.pulang.location
                    ]);
                }
                else {
                    worksheet.addRow([
                        item.nip,
                        item.nama,
                        new Date(item.datang.waktu * 1000),
                        item.datang.latlong,
                        item.datang.location
                    ]);
                }
                
                worksheet.getRow(i+3).commit();
            }

            workbook.xlsx.writeBuffer( {
                base64: true
            })
            .then( function (xls64) {
                // build anchor tag and attach file (works in chrome)
                var a = document.createElement("a");
                var data = new Blob([xls64], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

                var url = URL.createObjectURL(data);
                a.href = url;
                a.download = moment().format('YYYYMMDD-HHmmss') + ".xlsx";
                document.body.appendChild(a);
                a.click();
                setTimeout(function() {
                        document.body.removeChild(a);
                        window.URL.revokeObjectURL(url);
                    },
                    0);
            })
            .catch(function(error) {
                console.log(error.message);
            });
            // var ua = window.navigator.userAgent;
            // var msie = ua.indexOf("MSIE "); 
            // var tab_text="<table border='2px'><tr bgcolor='#87AFC6'>";
            // var textRange; var j=0;
            // tab = document.getElementById('tabel-absensi'); // id of table

            // for(j = 0 ; j < tab.rows.length ; j++) 
            // {     
            //     tab_text=tab_text+tab.rows[j].innerHTML+"</tr>";
            // }

            // tab_text=tab_text+"</table>";
            // tab_text= tab_text.replace(/<A[^>]*>|<\/A>/g, "");//remove if u want links in your table
            // tab_text= tab_text.replace(/<img[^>]*>/gi,""); // remove if u want images in your table
            // tab_text= tab_text.replace(/<input[^>]*>|<\/input>/gi, ""); // reomves input params

            // var ua = window.navigator.userAgent;
            // var msie = ua.indexOf("MSIE "); 
            // const tab_text = '<table>' + $('#tabel-absensi').html() + '</table>';
            // console.log(tab_text)

            // if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./))      // If Internet Explorer
            // {
            //     txtArea1.document.open("txt/html","replace");
            //     txtArea1.document.write(tab_text);
            //     txtArea1.document.close();
            //     txtArea1.focus(); 
            //     sa=txtArea1.document.execCommand("SaveAs",true,"Say Thanks to Sumit.xls");
            // }  
            // else                 //other browser not tested on IE 11
            //     sa = window.open('data:application/vnd.ms-excel,' + encodeURIComponent(tab_text));  

            // return (sa);
        }
    }
})