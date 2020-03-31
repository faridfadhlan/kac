const app = new Vue({
    el: '#app',
    data: {
        file: null,
        inputKaizala: [],
        absens: [],
        timezone: "",
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
            if(app.file == null) return Swal.fire('Error', 'File belum dipilih', 'error');
            // if(app.timezone == "" || app.timezone == null) return Swal.fire('Error', 'Timezone belum dipilih', 'error');
            Papa.parse(app.file, {
                header: true,
                error: (err) => {
                    if(err) {
                        return Swal.fire('Error', err.toString(), 'error')
                    }
                },
                complete: (results, f) => {
                    app.inputKaizala = results.data;
                    app.addUnixTime();
                    // return;
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
                d.unixservertime = moment(d['Server Receipt Timestamp (UTC)'], 'YYYY-MM-DD HH:mm:ss A').unix();
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
                    groupname: unik[k][0]['Group Name'],
                    datang: unik[k][0],
                    pulang: unik[k].length > 1 ? unik[k][unik[k].length-1] : null
                }
                return absens;
            });

            

            
            app.absens = absen_per_orang.map(apo => {
                // console.log(apo)
                const offset = new Date().getTimezoneOffset();
                const data = {};
                const pecahnama = apo.nama.split('-');
                data.nama = pecahnama[0].trim();
                data.nip = pecahnama[1].trim();
                const groupname = apo.groupname.split('-');
                data.kode_uk = groupname[0].trim();
                data.uk = groupname[1].trim();
                data.groupname = apo.groupname;
                data.datang = {
                    waktu: moment.unix(apo.datang.unixservertime).add((0-(offset/60)), 'h').toDate(),
                    latitude: apo.datang['Responder Location Latitude'],
                    longitude: apo.datang['Responder Location Longitude'],
                    location: apo.datang['Responder Location Location'],
                    latlong: apo.datang['Responder Location Latitude'] + ', ' + apo.datang['Responder Location Longitude']
                };
                if(apo.pulang) {
                    data.pulang = {
                        waktu: moment.unix(apo.pulang.unixservertime).add((0-(offset/60)), 'h').toDate(),
                        latitude: apo.pulang['Responder Location Latitude'],
                        longitude: apo.pulang['Responder Location Longitude'],
                        latlong: apo.pulang['Responder Location Latitude'] + ', ' + apo.pulang['Responder Location Longitude'],
                        location: apo.pulang['Responder Location Location']
                    }
                }
                else {
                    data.pulang = null
                }
                // console.log()
                
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
            worksheet.getColumn(3).width = 45;
            worksheet.getColumn(4).width = 20;
            worksheet.getColumn(5).width = 25;
            worksheet.getColumn(6).width = 40;
            worksheet.getColumn(7).width = 20;
            worksheet.getColumn(8).width = 25;
            worksheet.getColumn(9).width = 40;

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

            //
            worksheet.mergeCells('C1:C2');
            worksheet.getCell('C1').value = 'Nama Group';
            worksheet.getCell('C1').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('C1').font = { bold: true };

            // kolom absen datang
            worksheet.mergeCells('D1:F1');
            worksheet.getCell('D1').value = 'Absen Datang';
            worksheet.getCell('D1').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('D1').font = { bold: true };

            // kolom absen pulang
            worksheet.mergeCells('G1:I1');
            worksheet.getCell('G1').value = 'Absen Pulang'
            worksheet.getCell('G1').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('G1').font = { bold: true };

            worksheet.getCell('D2').value = 'Tanggal';
            worksheet.getCell('D2').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('D2').font = { bold: true };
            worksheet.getCell('E2').value = 'Latlong';
            worksheet.getCell('E2').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('E2').font = { bold: true };
            worksheet.getCell('F2').value = 'Alamat';
            worksheet.getCell('F2').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('F2').font = { bold: true };

            worksheet.getCell('G2').value = 'Tanggal';
            worksheet.getCell('G2').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('G2').font = { bold: true };
            worksheet.getCell('H2').value = 'Latlong';
            worksheet.getCell('H2').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('H2').font = { bold: true };
            worksheet.getCell('I2').value = 'Alamat';
            worksheet.getCell('I2').alignment = { vertical: 'center', horizontal: 'center' };
            worksheet.getCell('I2').font = { bold: true };
            

            worksheet.getRow(2).commit(); // now rows 1 and two are committed.
            // console.log(app.absens);return;

            for(i in app.absens) {
                const item = app.absens[i];
                if(item.pulang) {
                    worksheet.addRow([
                        item.nip,
                        item.nama,
                        item.groupname,
                        item.datang.waktu,
                        item.datang.latlong,
                        item.datang.location,
                        item.pulang.waktu,
                        item.pulang.latlong,
                        item.pulang.location
                    ]);
                }
                else {
                    // console.log('ddd')
                    worksheet.addRow([
                        item.nip,
                        item.nama,
                        item.groupname,
                        item.datang.waktu,
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
        }
    }
})