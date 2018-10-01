(function () {
    "use strict";
    var ClockType = Object.freeze({ "BAR_CLOCK": "bar", "ROUND_CLOCK": "round", "DIGITAL_CLOCK": "digital" });
    var messageBanner;
    var ctx;
    var radius;
    var timer;
    var canvasHeight, canvasWidth;
    var startTime = new Date();
    var isTimerStarted = false;
    var isTimeUp = false;
    var clockType = ClockType.ROUND_CLOCK;

    var HH = 0, MM = 0, SS = 20;
    var interval = 5;

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {           
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            for (let i = 0; i < 24; i++) {
                $("#toolbar-HH").append($('<option>', { value: i, text: i < 10 ? "0" + i.toString() : i.toString() }));
            }
            for (let i = 0; i < 60; i++) {
                $("#toolbar-MM").append($('<option>', { value: i, text: i < 10 ? "0" + i.toString() : i.toString() }));
            }
            for (let i = 0; i < 60; i++) {
                $("#toolbar-SS").append($('<option>', { value: i, text: i < 10 ? "0" + i.toString() : i.toString() }));
            }
            var canvas = document.getElementById("canvas");
            canvasHeight = canvas.height;
            canvasWidth = canvas.width;
            ctx = canvas.getContext("2d");
            radius = canvas.height / 2;
            ctx.translate(radius, radius);
            radius = radius * 0.9;

            drawRoundClock();

            window.addEventListener('resize', handleWindowResize);

            $("#toolbar-HH").val("0").trigger("change");
            $("#toolbar-MM").val("0").trigger("change");
            $("#toolbar-SS").val("20").trigger("change");
            $('#toolbar-interval').val("5").trigger("change");

            $('#toolbar-clocktype').on('change', handleClockTypeChange);
            $('#toolbar-HH').on('change', handleToolbarHHChange);
            $('#toolbar-MM').on('change', handleToolbarMMChange);
            $('#toolbar-SS').on('change', handleToolbarSSChange);
            $('#toolbar-interval').on('change', handleIntervalInputChange);

            $('#clock-start-btn').on('click', handleClockStartBtnPressed);
            $('#clock-stop-btn').on('click', handleClockStopBtnPressed);
            $('#content-main').on('click', toggleClockStatus);
          
            getActiveFileView();
            registerActiveViewChanged();
        });
    };

    function handleWindowResize() {
        $('body').css('height', window.innerHeight * 0.96);
       // $('#content-main').css('height', window.innerHeight * 0.9);
       // $('#content-tool').css('height', window.innerHeight * 0.9);
        var canvas_size = Math.min.apply(null, [$('#content-main').width() * 0.96, $('#content-main').height() * 0.96]);
       // showNotification("canvas size is " + canvas_size + "height is " + $('#content-main').width() + "width is " + $('#content-main').height());
        $('#canvas').css({ 'height': canvas_size, 'width': canvas_size});
    }

    function toggleClockStatus() {
        isTimerStarted = !isTimerStarted;
        if (isTimerStarted) {
            isTimeUp = false;
            handleClockStartBtnPressed();
        } else {
            handleClockStopBtnPressed();
        }
    }
    function handleClockStartBtnPressed() { 
        startTime = new Date();
        isTimerStarted = true;
        isTimeUp = false;
        if (clockType == ClockType.ROUND_CLOCK) {
            timer = setInterval(drawRoundClock, 1000);
        } else if (clockType == ClockType.DIGITAL_CLOCK) {
            timer = setInterval(drawDigitalClock, 1000);
        }
    }
    function handleClockStopBtnPressed() {
        isTimerStarted = false;
        clearInterval(timer);
    }
    function handleToolbarHHChange(e) {
        if (isTimerStarted) {
            isTimerStarted = false;
        }
        HH = parseInt(this.value);
        isTimerStarted = false;
        drawClock();
    }
    function handleToolbarMMChange(e) {
        if (isTimerStarted) {
            isTimerStarted = false;
        }
        MM = parseInt(this.value);
        isTimerStarted = false;
        drawClock();
    }
    function handleToolbarSSChange() {
        if (isTimerStarted) {
            isTimerStarted = false;
        }
        SS = parseInt(this.value);
        isTimerStarted = false;
        drawClock();
    }
    function handleIntervalInputChange() {
        if (isTimerStarted) {
            isTimerStarted = false;
        }
        interval = parseInt(this.value);
        isTimerStarted = false;
        drawClock();
    }
    function drawClock() {
        switch (clockType) {
            case ClockType.BAR_CLOCK:
                break;
            case ClockType.DIGITAL_CLOCK:
                drawDigitalClock();
                break;
            case ClockType.ROUND_CLOCK:
                drawRoundClock();
                break;
        }
    }
    function handleClockTypeChange() {
        switch (this.value) {
            case ClockType.BAR_CLOCK:
                clockType = ClockType.BAR_CLOCK;
                break;
            case ClockType.ROUND_CLOCK:
                clockType = ClockType.ROUND_CLOCK;
                ctx.translate(canvasHeight / 2, canvasHeight / 2);
                ctx.clearRect(0, 0, canvasWidth, canvasHeight);
                drawRoundClock();
                break;
            case ClockType.DIGITAL_CLOCK:
                clockType = ClockType.DIGITAL_CLOCK;
                ctx.translate(-canvasHeight / 2, -canvasHeight/2);
                ctx.clearRect(0, 0, canvasWidth, canvasHeight);
                drawDigitalClock();
                break;
        }
    }

    function getActiveFileView() {
        Office.context.document.getActiveViewAsync(function (asyncResult) {
            if (asyncResult.status == "failed") {
                showNotification("Action failed with error: " + asyncResult.error.message);
            }
            else {
                var toolbar = document.getElementById("content-tool");
                var clockarea = document.getElementById("content-main");
                if (asyncResult.value == "read") {
                    toolbar.style.display = "none";
                    clockarea.style.width = "100%";
                    var canvas_size = Math.min.apply(null, [$('#content-main').width() * 0.96, $('#content-main').height() * 0.96]);
                    $('#canvas').css({ 'height': canvas_size, 'width': canvas_size });
                    showNotification("canvas size is " + canvas_size + "height is " + $('#content-main').width() + "width is " + $('#content-main').height());
                } else {
                    toolbar.style.display = "block";
                    clockarea.style.width = "60%";
                }
            }
        });
    }

    function registerActiveViewChanged() {
        Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged,
            getActiveFileView);
    }


    function drawRoundClock() {
        drawClockFace(ctx, radius);
        drawClockNumbers(ctx, radius);
        drawClockTime(ctx, radius);
    }

    function drawClockFace(ctx, radius) {
        var grad;

        ctx.beginPath();
        ctx.arc(0, 0, radius, 0, 2 * Math.PI);
        ctx.fillStyle = 'white';
        ctx.fill();

        grad = ctx.createRadialGradient(0, 0, radius * 0.95, 0, 0, radius * 1.05);
        grad.addColorStop(0, '#333');
        grad.addColorStop(0.5, 'white');
        grad.addColorStop(1, '#333');
        ctx.strokeStyle = grad;
        ctx.lineWidth = radius * 0.1;
        ctx.stroke();

        ctx.beginPath();
        ctx.arc(0, 0, radius * 0.1, 0, 2 * Math.PI);
        ctx.fillStyle = '#333';
        ctx.fill();
    }

    function drawClockNumbers(ctx, radius) {
        var ang;
        var numIntervals = (HH * 3600 + MM * 60 + SS) / interval;
        ctx.font = radius * 0.15 + "px arial";
        ctx.textBaseline = "middle";
        ctx.textAlign = "center";
        for (var num = 1; num < numIntervals + 1; num++) {
            ang = num * 2 * Math.PI / numIntervals;
            ctx.rotate(ang);
            ctx.translate(0, -radius * 0.85);
            ctx.rotate(-ang);
            ctx.fillText(".", 0, 0);
            ctx.rotate(ang);
            ctx.translate(0, radius * 0.85);
            ctx.rotate(-ang);
        }
    }

    function drawClockTime(ctx, radius) {
        var totalDuration = HH * 3600 + MM * 60 + SS;
        var startTimeInSec = startTime.getHours() * 3600 + startTime.getMinutes() * 60 + startTime.getSeconds();
        var now = new Date();
        var currTimeInSec = now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds();
        var radToRotate = (currTimeInSec - startTimeInSec) * 1.0 / totalDuration * 2 * Math.PI;
        if (currTimeInSec - startTimeInSec > totalDuration || !isTimerStarted) {
            clearInterval(timer);
            isTimerStarted = false;
            drawClockHand(ctx, 0, radius * 0.9, radius * 0.02); //reset
        } else {
            drawClockHand(ctx, radToRotate, radius * 0.9, radius * 0.02);
        }
    }

    function drawClockHand(ctx, pos, length, width) {
        ctx.beginPath();
        ctx.lineWidth = width;
        ctx.lineCap = "round";
        ctx.moveTo(0, 0);
        ctx.rotate(pos);
        ctx.lineTo(0, -length);
        ctx.stroke();
        ctx.rotate(-pos);
    }

    function drawDigitalClock() {
        ctx.clearRect(0, 0, canvasWidth, canvasHeight);
        ctx.font = "80pt calibri";
        ctx.fillStyle = "black";
        if (isTimeUp) {
            clearInterval(timer);
            ctx.fillText(formatHHMMSS(0, 0, 0), canvasWidth / 2, canvasHeight / 2);
        } else if (isTimerStarted) {
            ctx.fillText(calculateDigitalTime(), canvasWidth / 2, canvasHeight / 2);
        } else {
            ctx.fillText(formatHHMMSS(HH, MM, SS), canvasWidth / 2, canvasHeight / 2);
        }
    }

    function calculateDigitalTime() {
        var totalDuration = HH * 3600 + MM * 60 + SS;
        var startTimeInSec = startTime.getHours() * 3600 + startTime.getMinutes() * 60 + startTime.getSeconds();
        var now = new Date();
        var currTimeInSec = now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds();
        var timeLeft = totalDuration - (currTimeInSec - startTimeInSec);
        var hourLeft = 0, minLeft = 0, ssLeft = 0;
        if (timeLeft > 0) {
            hourLeft = Math.floor(timeLeft / 3600);
            minLeft = Math.floor((timeLeft % 3600) / 60);
            ssLeft = timeLeft - hourLeft * 3600 - minLeft * 60;
        } else {
            isTimeUp = true;
            isTimerStarted = false;
        }
        var rltStr = formatHHMMSS(hourLeft, minLeft, ssLeft);
        return rltStr;
    }

    function formatHHMMSS(hh, mm, ss) {
        var rltStr = "";
        if (hh < 10) {
            rltStr += "0" + hh.toString() + " : ";
        } else {
            rltStr += hh.toString() + " : ";
        }
        if (mm < 10) {
            rltStr += "0" + mm.toString() + " : ";
        } else {
            rltStr += mm.toString() + " : ";
        }
        if (ss < 10) {
            rltStr += "0" + ss.toString();
        } else {
            rltStr += ss.toString();
        }
        return rltStr;
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();