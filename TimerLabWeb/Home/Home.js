(function () {
    "use strict";
    var ClockType = Object.freeze({ "BAR_CLOCK": "bar", "SQUARE_CLOCK": "square", "DIGITAL_CLOCK": "digital" });
    var clockType = ClockType.SQUARE_CLOCK;
    var messageBanner;
    var ctx;
    var radius;
    var timer;
    var canvasHeight, canvasWidth;
    var startTime = new Date();
    var isTimerStarted = false;
    var isTimeUp = false;
    var isReset = false;
    var barPtrx, barPtry;

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
        isTimeUp = false;
        clearInterval(timer);
    }
    function handleToolbarHHChange(e) {
        if (isTimerStarted) {
            isTimerStarted = false;
        }
        if ((parseInt(this.value) * 3600 + MM * 60 + SS) % interval != 0
            && clockType != ClockType.DIGITAL_CLOCK) {
            $("#error-message-duration").text("interval value must be divisible by duration");
            return;
        }
        if (!isReset) {
            $("#error-message-duration").text("hour changed successfully!");
            setTimeout(function () {
                $("#error-message-duration").text("");
            }, 2000);
        } else {
            $("#error-message-duration").text("");
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
        if ((HH * 3600 + parseInt(this.value) * 60 + SS) % interval != 0
            && clockType != ClockType.DIGITAL_CLOCK) {
            $("#error-message-duration").text("interval value must be divisible by duration");
            return;
        }
        if (!isReset) {
            $("#error-message-duration").text("minute changed successfully!");
            setTimeout(function () {
                $("#error-message-duration").text("");
            }, 2000);
        } else {
            $("#error-message-duration").text("");
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
        if ((HH * 3600 + MM * 60 + parseInt(this.value)) % interval != 0
            && clockType != ClockType.DIGITAL_CLOCK) {
            $("#error-message-duration").text("interval value must be divisible by duration");
            return;
        }
        if (!isReset) {
            $("#error-message-duration").text("second changed successfully!");
            setTimeout(function () {
                $("#error-message-duration").text("");
            }, 2000);
        } else {
            $("#error-message-duration").text("");
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
        if (!/^\s*\d+\s*$/.test(this.value)) {
            $("#error-message").text("interval value must be integer");
            return;
        } else if (parseInt(this.value) == 0) {
            $("#error-message").text("interval value cannot be zero");
            return;
        } else if (parseInt(this.value) > 60) {
            $("#error-message").text("interval value should be less than 60");
            return;
        } else if ((HH * 3600 + MM * 60 + SS) % parseInt(this.value) != 0) {
            $("#error-message").text("interval value must be divisible by duration");
            return;
        }
        if (!isReset) {
            $("#error-message").text("interval changed successfully!");
            setTimeout(function () {
                $("#error-message").text("");
            }, 2000);
        } else {
            $("#error-message").text("");
        }
        interval = parseInt(this.value);
        saveSettings('interval', interval);
        isTimerStarted = false;
        drawClock();
    }
    
    function handleClockTypeChange() {
        saveSettings('clocktype', this.value);
        reset();
        switch (this.value) {
            case ClockType.BAR_CLOCK:
                clockType = ClockType.BAR_CLOCK;
                $("#interval-div").css("display", "block");
                break;
            case ClockType.SQUARE_CLOCK:
                clockType = ClockType.SQUARE_CLOCK;
                $("#interval-div").css("display", "block");
                break;
            case ClockType.DIGITAL_CLOCK:
                clockType = ClockType.DIGITAL_CLOCK;
                $("#interval-div").css("display", "none");
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

    // #region SquareClock

    function drawSquareClock() {
        drawSqClockFace(ctx, radius);
        drawSqClockNumbers(ctx, radius);
        drawSqClockTime(ctx, radius);
    }

    function drawSqClockFace(ctx, radius) {
        var grad;

        ctx.beginPath();
        ctx.rect(-radius, -radius, radius * 2, radius * 2);
        ctx.fillStyle = 'white';
        ctx.fill();

        grad = ctx.createLinearGradient(-radius * 0.95, -radius * 0.95, radius * 1.05, radius * 1.05);
        for (var i = 0; i < 1; i += 0.2) {
            grad.addColorStop(i, '#333');
            grad.addColorStop(i + 0.1, 'white');
        }
        grad.addColorStop(1, '#333');
        ctx.strokeStyle = grad;
        ctx.lineWidth = radius * 0.1;
        ctx.stroke();

        ctx.beginPath();
        ctx.rect(-radius * 0.1, -radius * 0.1, radius * 0.2, radius * 0.2);
        ctx.fillStyle = '#333';
        ctx.fill();
    }

    function drawSqClockNumbers(ctx, radius) {
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

    function drawSqClockTime(ctx, radius) {
        var totalDuration = HH * 3600 + MM * 60 + SS;
        var startTimeInSec = startTime.getHours() * 3600 + startTime.getMinutes() * 60 + startTime.getSeconds();
        var now = new Date();
        var currTimeInSec = now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds();
        var radToRotate = (currTimeInSec - startTimeInSec) * 1.0 / totalDuration * 2 * Math.PI;
        if (currTimeInSec - startTimeInSec > totalDuration || !isTimerStarted) {
            clearInterval(timer);
            isTimerStarted = false;
            drawSqClockHand(ctx, 0, radius * 0.9, radius * 0.02); //reset
        } else {
            drawSqClockHand(ctx, radToRotate, radius * 0.9, radius * 0.02);
        }
    }

    function drawSqClockHand(ctx, pos, length, width) {
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
            ctx.fillText(formatHHMMSS(0, 0, 0), canvasWidth/2, canvasHeight / 2);
        } else if (isTimerStarted) {
            ctx.fillText(calculateDigitalTime(), canvasWidth/2, canvasHeight / 2);
        } else {
            ctx.fillText(formatHHMMSS(HH, MM, SS), canvasWidth/2, canvasHeight / 2);
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
        var rectw = canvasWidth * 0.8;
        var recth = canvasHeight / 5;
        drawBarClockBody(rectw, recth);
        drawBarClockNumber(rectw, recth);
        drawBarClockPointer(rectw, recth);
    }

    function drawBarClockBody(rectw, recth) {   
        barPtrx = canvasWidth / 2 - rectw / 2;
        barPtry = canvasHeight / 2 - recth / 2;
        ctx.beginPath();
        ctx.rect(canvasWidth / 2 - rectw / 2, canvasHeight / 2 - recth / 2, rectw, recth);
        ctx.fillStyle = "#104B10";
        ctx.fill();
    }
    function drawBarClockPointer(rectw, recth) {
        if (isTimerStarted) {
            var totalDuration = HH * 3600 + MM * 60 + SS;
            var startTimeInSec = startTime.getHours() * 3600 + startTime.getMinutes() * 60 + startTime.getSeconds();
            var now = new Date();
            var currTimeInSec = now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds();
            if (currTimeInSec - startTimeInSec > totalDuration) {
                isTimeUp = true;
                isTimerStarted = false;
                barPtrx = canvasWidth / 2 + rectw / 2;
                clearInterval(timer);
            } else {
                var deltaw = (currTimeInSec - startTimeInSec) / totalDuration * rectw;
                barPtrx += deltaw;
            }
        }
        if (isTimeUp) {
            barPtrx = canvasWidth / 2 + rectw / 2;
            clearInterval(timer);
        }
        if (!isTimeUp && !isTimerStarted) {
            barPtrx = canvasWidth / 2 - rectw / 2;
        }
        var rectPtrw = canvasWidth * 0.01;
        var rectPtrh = recth;
        ctx.beginPath();
        ctx.rect(barPtrx, barPtry, rectPtrw, rectPtrh);
        ctx.fillStyle = "#F39F16";
        ctx.fill();

        var triPtrw = canvasWidth * 0.1;
        var triPtrh = canvasWidth * 0.05;
        ctx.beginPath();
        ctx.moveTo(barPtrx - triPtrw / 2, barPtry - triPtrh + rectPtrw / 2);
        ctx.lineTo(barPtrx + triPtrw / 2, barPtry - triPtrh + rectPtrw / 2);
        ctx.lineTo(barPtrx + rectPtrw / 2, barPtry + rectPtrw / 2);
        ctx.fillStyle = "#F39F16";
        ctx.fill();
    }
    function drawBarClockNumber(rectw, recth) {
        var lbx = barPtrx;
        var lby = barPtry + recth;
        var numIntervals = (HH * 3600 + MM * 60 + SS) / interval;
        var intervalw = rectw / numIntervals;
        var intervalh = recth / 10;
        for (var i = 0; i <= numIntervals; i++) {
            ctx.beginPath();
            ctx.moveTo(lbx + i * intervalw, lby);
            ctx.lineTo(lbx + i * intervalw, lby - intervalh);
            ctx.strokeStyle = "black";
            ctx.stroke();
        }
    }

    // #endregion

    // #region OtherFunctions

    function drawClock() {
        var canvas = document.getElementById("canvas");
        canvasHeight = canvas.height;
        canvasWidth = canvas.width;
        ctx = canvas.getContext("2d");
        ctx.clearRect(0, 0, canvasWidth, canvasHeight);
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        switch (clockType) {
            case ClockType.BAR_CLOCK:
                drawBarClock();
                break;
            case ClockType.DIGITAL_CLOCK:
                drawDigitalClock();
                break;
            case ClockType.SQUARE_CLOCK:
                var size = Math.min.apply(null, [canvasHeight, canvasWidth]) / 2;
                radius = size * 0.9;
                ctx.translate(size, size);
                drawSquareClock();
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

    function reset() {
            isReset = true;
            HH = 0;
            MM = 0;
            SS = 20;
            interval = 5;
            $("#toolbar-HH").val(HH).trigger("change");
            $("#toolbar-MM").val(MM).trigger("change");
            $("#toolbar-SS").val(SS).trigger("change");
            $('#toolbar-interval').val(interval).trigger("change");
            saveSettings('HH', HH);
            saveSettings('MM', MM);
            saveSettings('SS', SS);
        saveSettings('interval', interval);
        isReset = false;
    }

    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    // #endregion
})();