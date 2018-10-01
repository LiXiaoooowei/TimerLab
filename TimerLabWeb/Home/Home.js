﻿(function () {
    "use strict";
    var ClockType = Object.freeze({ "BAR_CLOCK": "bar", "ROUND_CLOCK": "round", "DIGITAL_CLOCK": "digital" });
    var clockType = ClockType.ROUND_CLOCK;
    var messageBanner;
    var ctx;
    var radius;
    var timer;
    var canvasHeight, canvasWidth;
    var startTime = new Date();
    var isTimerStarted = false;
    var isTimeUp = false;

    var HH = 0, MM = 0, SS = 20;
    var interval = 5;

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {           
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            handleActiveFileViewChanged();
            handleActiveViewChanged();

            for (let i = 0; i < 24; i++) {
                $("#toolbar-HH").append($('<option>', { value: i, text: i < 10 ? "0" + i.toString() : i.toString() }));
            }
            for (let i = 0; i < 60; i++) {
                $("#toolbar-MM").append($('<option>', { value: i, text: i < 10 ? "0" + i.toString() : i.toString() }));
            }
            for (let i = 0; i < 60; i++) {
                $("#toolbar-SS").append($('<option>', { value: i, text: i < 10 ? "0" + i.toString() : i.toString() }));
            }

            HH = loadSettings('HH') == null ? HH : loadSettings('HH');
            MM = loadSettings('MM') == null ? MM : loadSettings('MM');
            SS = loadSettings('SS') == null ? SS : loadSettings('SS');
            interval = loadSettings('interval') == null ? interval : loadSettings('interval');
            clockType = loadSettings('clocktype') == null ? clockType : loadSettings('clocktype');

            var canvas = document.getElementById("canvas");
            canvasHeight = canvas.height;
            canvasWidth = canvas.width;
            ctx = canvas.getContext("2d");
            
            drawClock();
            
            window.addEventListener('resize', handleWindowResize);
   
            $("#toolbar-HH").val(HH).trigger("change");
            $("#toolbar-MM").val(MM).trigger("change");
            $("#toolbar-SS").val(SS).trigger("change");
            $('#toolbar-interval').val(interval).trigger("change");
            $('#toolbar-clocktype').val(clockType).trigger("change");

            $('#toolbar-clocktype').on('change', handleClockTypeChange);
            $('#toolbar-HH').on('change', handleToolbarHHChange);
            $('#toolbar-MM').on('change', handleToolbarMMChange);
            $('#toolbar-SS').on('change', handleToolbarSSChange);
            $('#toolbar-interval').on('change', handleIntervalInputChange);

            $('#clock-start-btn').on('click', handleClockStartBtnPressed);
            $('#clock-stop-btn').on('click', handleClockStopBtnPressed);
            $('#content-main').on('click', handleClockStatusChanged);
     
        });
    };
    // #region EventHandlers
    function handleWindowResize() {
        $('body').css('height', window.innerHeight * 0.96);
        var size = Math.min.apply(null, [$('#content-main').height() * 0.96, $('#content-main').width() * 0.96]);
        $('#canvas').css({ 'height': size, 'width': size});
        canvasWidth = $('#canvas').width();
        canvasHeight = $('#canvas').height();
        radius = 0.9 * Math.min.apply(null, [canvasHeight, canvasWidth]) / 2;
    }

    function handleClockStatusChanged() {
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
        timer = setInterval(drawClock, 1000);
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
        saveSettings('HH', HH);
        isTimerStarted = false;
        drawClock();
    }
    function handleToolbarMMChange(e) {
        if (isTimerStarted) {
            isTimerStarted = false;
        }
        MM = parseInt(this.value);
        saveSettings('MM', MM);
        isTimerStarted = false;
        drawClock();
    }
    function handleToolbarSSChange() {
        if (isTimerStarted) {
            isTimerStarted = false;
        }
        SS = parseInt(this.value);
        saveSettings('SS', SS);
        isTimerStarted = false;
        drawClock();
    }
    function handleIntervalInputChange() {
        if (isTimerStarted) {
            isTimerStarted = false;
        }
        interval = parseInt(this.value);
        saveSettings('interval', interval);
        isTimerStarted = false;
        drawClock();
    }
    
    function handleClockTypeChange() {
        saveSettings('clocktype', this.value);
        switch (this.value) {
            case ClockType.BAR_CLOCK:
                clockType = ClockType.BAR_CLOCK;
                break;
            case ClockType.ROUND_CLOCK:
                clockType = ClockType.ROUND_CLOCK;
                break;
            case ClockType.DIGITAL_CLOCK:
                clockType = ClockType.DIGITAL_CLOCK;
                break;
        }
        drawClock();
    }

    function handleActiveViewChanged() {
        Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged,
            handleActiveFileViewChanged);
    }

    function handleActiveFileViewChanged() {
        Office.context.document.getActiveViewAsync(function (asyncResult) {
            if (asyncResult.status == "failed") {
                showNotification("Action failed with error: " + asyncResult.error.message);
            }
            else {
                var toolbar = document.getElementById("content-tool");
                var clockarea = document.getElementById("content-main");
                if (asyncResult.value == "read") {
                    toolbar.style.display = "none";
                    $("#content-main").css('margin-left', '20 %');
                    var size = Math.min.apply(null, [$('#content-main').height() * 0.96, $('#content-main').width() * 0.96]);
                    $('#canvas').css({ 'height': size, 'width': size});                   
                    canvasWidth = $('#canvas').width();
                    canvasHeight = $('#canvas').height();
                } else {
                    toolbar.style.display = "block";
                    clockarea.style.width = "60%";
                }
            }
        });
    }

    // #endregion

    // #region RoundClock

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

    // #endregion

    // #region DigitalClock

    function drawDigitalClock() {
        ctx.font = "80pt calibri";
        ctx.fillStyle = "black";
        ctx.textAlign = "center";
        if (isTimeUp) {
            clearInterval(timer);
            ctx.strokeText(formatHHMMSS(0, 0, 0), canvasWidth/2, canvasHeight / 2);
        } else if (isTimerStarted) {
            ctx.strokeText(calculateDigitalTime(), canvasWidth/2, canvasHeight / 2);
        } else {
            ctx.strokeText(formatHHMMSS(HH, MM, SS), canvasWidth/2, canvasHeight / 2);
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

    // #endregion

    // #region BarClock

    function drawBarClock() {
    }

    // #endregion

    // #region OtherFunctions

    function drawClock() {
        switch (clockType) {
            case ClockType.BAR_CLOCK:
                drawBarClock();
                break;
            case ClockType.DIGITAL_CLOCK:
                var canvas = document.getElementById("canvas");
                canvasHeight = canvas.height;
                canvasWidth = canvas.width;
                ctx = canvas.getContext("2d");
                ctx.clearRect(0, 0, canvasWidth, canvasHeight);
                ctx.setTransform(1, 0, 0, 1, 0, 0);
                drawDigitalClock();
                break;
            case ClockType.ROUND_CLOCK:
                var canvas = document.getElementById("canvas");
                canvasHeight = canvas.height;
                canvasWidth = canvas.width;
                ctx = canvas.getContext("2d");
                ctx.clearRect(0, 0, canvasWidth, canvasHeight);
                var size = Math.min.apply(null, [canvasHeight, canvasWidth]) / 2;
                radius = size * 0.9;
                ctx.translate(size, size);
                drawRoundClock();
                ctx.translate(-size, -size);
                break;
        }
    }

    function saveSettings(key, value) {
        Office.context.document.settings.set(key, value);
        Office.context.document.settings.saveAsync(function (asyncResult) {
            console.log('Settings saved with status: ' + asyncResult.status);
        });
    }

    function loadSettings(key) {
        return Office.context.document.settings.get(key);
    }

    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    // #endregion
})();