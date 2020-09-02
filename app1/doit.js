
            
                // Office is ready
                $(document).ready(function () {
                    console.log("test1");
                    //var element1 = document.getElementById('a');
                    Office.onReady(function() {
                    //element1.innerHTML = "initialize1";
                    //document.write("test");
                    console.log("test_xxxxxxxxxx");

                            var element1 = document.getElementById('a');
                            var element2 = document.getElementById('b');

                            element1.innerHTML = "initialize31";
                            var element1 = document.getElementById('a');
                            var element2 = document.getElementById('b');
                            element1.innerHTML = "EWS URL: " + Office.context.mailbox.ewsUrl;
                            element2.innerHTML = "REST URL: " + Office.context.mailbox.restUrl;
                            
                            console.log("REST URL: " + Office.context.mailbox.restUrl);
                            console.log("EWS URL: " + Office.context.mailbox.ewsUrl);



                });
            });