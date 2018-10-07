(function () {
    "use strict";

    var ClockType = Object.freeze({ "BAR_CLOCK": "bar", "SQUARE_CLOCK": "square", "DIGITAL_CLOCK": "digital" });
    var TickType = Object.freeze({ "NONE": "none", "TICK": "tick" });
    var TimeupType = Object.freeze({ "NONE": "none", "ALARM": "alarm" });

    var clockType = ClockType.SQUARE_CLOCK;
    var tickType = TickType.TICK;
    var timeupType = TimeupType.ALARM;

    var messageBanner;
    var ctx;
    var radius;
    var timer;
    var pausedTimer;
    var canvasHeight, canvasWidth;
    var startTime = null;
    var isTimerStarted = false;
    var hasInstructionToStartTimer = false;
    var isTimeUp = false;
    var isReset = false;
    var pausedTimeInSec = 0;
    var pausedDuration = 0;
    var isPaused = false;
    var showCombi = false;
    var isCountUp = false;
    var barPtrx, barPtry;

    var HH = 0, MM = 0, SS = 20;
    var interval = 5;

    var tick_audio = new Audio('../Resources/Audio/ticking.mp3');
    var timeup_audio = new Audio('../Resources/Audio/alarm.mp3');

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
            tickType = loadSettings('tickType') == null ? tickType : loadSettings('tickType');
            timeupType = loadSettings('timeupType') == null ? timeupType : loadSettings('timeupType');
            showCombi = loadSettings('showCombi') == null ? false : loadSettings('showCombi');
            isCountUp = loadSettings('isCountUp') == null ? false : loadSettings('isCountUp');
            console.log("value is " + loadSettings('showCombi') + " " + loadSettings('isCountUp'));
            loadTickSound();
            loadTimeupSound();

            var canvas = document.getElementById("canvas");
            canvasHeight = loadSettings('canvash') == null ? canvas.height : loadSettings('canvash');
            canvasWidth = loadSettings('canvasw') == null ? canvas.width : loadSettings('canvasw');
            radius = loadSettings('radius') == null ? 0.9 * Math.min.apply(null, [canvasHeight, canvasWidth]) / 2 : loadSettings('radius');
            $('#canvas').css({ 'height': canvasHeight, 'width': canvasWidth });
            ctx = canvas.getContext("2d");

            window.addEventListener('resize', handleWindowResize);

         //   drawClock();

            $('#toolbar-clocktype').on('change', handleClockTypeChange);
            $('#toolbar-HH').on('change', handleToolbarHHChange);
            $('#toolbar-MM').on('change', handleToolbarMMChange);
            $('#toolbar-SS').on('change', handleToolbarSSChange);
            $('#toolbar-interval').on('change', handleIntervalInputChange);

            $('#clock-start-btn').on('click', handleClockStartBtnPressed);
            $('#clock-stop-btn').on('click', handleClockStopBtnPressed);
            $('#clock-pause-btn').on('click', handleClockPauseBtnPressed);
            $('#clock-reset-btn').on('click', handleClockResetBtnPressed);
            $('#content-main').on('click', handleClockStatusChanged);

            $("#toolbar-ticking-sound").on('change', handleTickSoundChange);
            $("#toolbar-timeup-sound").on('change', handleTimeupSoundChange);

            $("#checkbox-digital-bar-clock").on('change', handleClockCombiChange);
            $("#checkbox-digital-count-up").on('change', handleDigiClockCountUpChange);

            $('#toolbar-clocktype').val(clockType).trigger("change");
            $("#toolbar-HH").val(HH).trigger("change");
            $("#toolbar-MM").val(MM).trigger("change");
            $("#toolbar-SS").val(SS).trigger("change");
            $('#toolbar-interval').val(interval).trigger("change");
            $("#toolbar-ticking-sound").val(tickType).trigger("change");
            $("#toolbar-timeup-sound").val(timeupType).trigger("change");

            if (showCombi) {
                $("#checkbox-digital-bar-clock").prop("checked", true).trigger('change');
            }

            if (isCountUp && clockType == ClockType.DIGITAL_CLOCK) {
                $("#checkbox-digital-count-up").prop("checked", true).trigger('change');
           }
        });
    };
    // #region EventHandlers
    function handleWindowResize() {
        $('body').css('height', window.innerHeight * 0.96);
        var size = Math.min.apply(null, [$('#content-main').height() * 0.96, $('#content-main').width() * 0.96]);
        $('#canvas').css({ 'height': size, 'width': size });
        canvasWidth = $('#canvas').width();
        canvasHeight = $('#canvas').height();
        radius = 0.9 * Math.min.apply(null, [canvasHeight, canvasWidth]) / 2;
        saveSettings('canvasw', canvasWidth);
        saveSettings('canvash', canvasHeight);
        saveSettings('radius', radius);
    }

    function handleClockStatusChanged() {
        hasInstructionToStartTimer = !hasInstructionToStartTimer;
        if (hasInstructionToStartTimer) {
            handleClockStartBtnPressed();
        } else {
            handleClockStopBtnPressed();
        }
    }
    function handleClockStartBtnPressed() {
        if (!isPaused && !isTimerStarted) {
            startTime = new Date();
            isTimerStarted = true;
            isTimeUp = false;
            isPaused = false;
            pausedDuration = pausedDuration == 0 ? 0 : pausedDuration + 1;
            clearInterval(pausedTimer);
            timer = setInterval(drawClock, 1000);
        } else if (isPaused && isTimerStarted) {
            isTimerStarted = true;
            isTimeUp = false;
            isPaused = false;
            pausedDuration = pausedDuration == 0 ? 0 : pausedDuration + 1;
            clearInterval(pausedTimer);
            timer = setInterval(drawClock, 1000);
        }
    }
    function handleClockStopBtnPressed() {
        isTimerStarted = false;
        isTimeUp = false;
        pausedTimeInSec = 0;
        pausedDuration = 0;
        isPaused = false;
        clearInterval(timer);
        clearInterval(pausedTimer);
        drawClock();
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
        // drawClock();
        handleClockStopBtnPressed();
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
        // drawClock();
        handleClockStopBtnPressed();
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
        // drawClock();
        handleClockStopBtnPressed();
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
        // drawClock();
        handleClockStopBtnPressed();
    }

    function handleClockTypeChange() {
        saveSettings('clocktype', this.value);
      //  reset();
        switch (this.value) {
            case ClockType.BAR_CLOCK:
                clockType = ClockType.BAR_CLOCK;
                $("#interval-div").css("display", "block");
                $("#checkbox-digital-bar-clock").css("display", "inline");
                $("#checkbox-digital-bar-text").css("display", "inline");
                $("#checkbox-digital-bar-text").text("Combine With Digital Clock");
                $("#checkbox-digital-count-up").css("display", "none");
                $("#checkbox-digital-count-up-text").css("display", "none");
                break;
            case ClockType.SQUARE_CLOCK:
                clockType = ClockType.SQUARE_CLOCK;
                $("#interval-div").css("display", "block");
                $("#checkbox-digital-bar-clock").css("display", "none");
                $("#checkbox-digital-bar-text").css("display", "none");
                $("#checkbox-digital-count-up").css("display", "none");
                $("#checkbox-digital-count-up-text").css("display", "none");
                break;
            case ClockType.DIGITAL_CLOCK:
                clockType = ClockType.DIGITAL_CLOCK;
                $("#interval-div").css("display", "none");
                $("#checkbox-digital-bar-clock").css("display", "inline");
                $("#checkbox-digital-bar-text").css("display", "inline");
                $("#checkbox-digital-bar-text").text("Combine With Bar Clock");
                $("#checkbox-digital-count-up").css("display", "inline");
                $("#checkbox-digital-count-up-text").css("display", "inline");
                break;
        }
        drawClock();
    }

    function handleClockCombiChange() {
        if (this.checked) {
            showCombi = true;
        } else {
            showCombi = false;
        }
        saveSettings('showCombi', showCombi);
        showNotification(loadSettings('showCombi'));
        drawClock();
    }

    function handleDigiClockCountUpChange() {
        if (this.checked) {
            isCountUp = true;
        } else {
            isCountUp = false;
        }
        saveSettings('isCountUp', isCountUp);
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
                    $('#canvas').css({ 'height': size, 'width': size });
                    canvasWidth = $('#canvas').width();
                    canvasHeight = $('#canvas').height();
                } else {
                    toolbar.style.display = "block";
                    clockarea.style.width = "60%";
                }
            }
        });
    }

    function handleTickSoundChange() {
        tickType = this.value;
        saveSettings('tickType', tickType);
        loadTickSound();
    }

    function handleTimeupSoundChange() {
        timeupType = this.value;
        saveSettings('timeupType', timeupType);
        loadTimeupSound();
    }

    function handleClockPauseBtnPressed() {
        if (isTimerStarted && !isPaused) {
            isPaused = true;
            pausedTimeInSec = new Date();
            pausedTimeInSec = pausedTimeInSec.getHours() * 3600 + pausedTimeInSec.getMinutes() * 60 + pausedTimeInSec.getSeconds();
            pausedTimer = setInterval(updatePauseDuration, 1000);
            clearInterval(timer);
        }
    }

    function handleClockResetBtnPressed() {
        reset();
        tickType = TickType.TICK;
        timeupType = TimeupType.ALARM;
        loadTickSound();
        loadTimeupSound();

        $("#toolbar-ticking-sound").val(tickType).trigger("change");
        $("#toolbar-timeup-sound").val(timeupType).trigger("change");

        saveSettings('tickType', tickType);
        saveSettings('timeupType', timeupType);
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
        var startTimeInSec = startTime == null ? 0 : startTime.getHours() * 3600 + startTime.getMinutes() * 60 + startTime.getSeconds() + pausedDuration;
        var now = new Date();
        var currTimeInSec = now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds();
        console.log("difference is " + (currTimeInSec - startTimeInSec));
        var radToRotate = (currTimeInSec - startTimeInSec) * 1.0 / totalDuration * 2 * Math.PI;
        if (currTimeInSec - startTimeInSec == totalDuration) {
            pausedDuration = 0;
            pausedTimeInSec = 0;
            isPaused = false;
            if (timeup_audio != null) {
                timeup_audio.play();
            }
        }
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

    function drawDigitalClock(posY) {
        ctx.font = "80pt calibri";
        ctx.fillStyle = "black";
        ctx.textAlign = "center";
        if (isTimeUp) {
            clearInterval(timer);
            if (isCountUp) {
                ctx.fillText(formatHHMMSS(HH, MM, SS), canvasWidth / 2, posY);
            } else {
                ctx.fillText(formatHHMMSS(0, 0, 0), canvasWidth / 2, posY);
            }
        } else if (isTimerStarted) {
            ctx.fillText(calculateDigitalTime(), canvasWidth / 2, posY);
        } else if (!isTimerStarted) {
            clearInterval(timer);
            if (isCountUp) {
                ctx.fillText(formatHHMMSS(0, 0, 0), canvasWidth / 2, posY);
            } else {
                ctx.fillText(formatHHMMSS(HH, MM, SS), canvasWidth / 2, posY);
            }
        }
    }

    function calculateDigitalTime() {
        var totalDuration = HH * 3600 + MM * 60 + SS;
        var startTimeInSec = startTime == null ? 0 : startTime.getHours() * 3600 + startTime.getMinutes() * 60 + startTime.getSeconds() + pausedDuration;
        var now = new Date();
        var currTimeInSec = now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds();
        var timeLeft = totalDuration - (currTimeInSec - startTimeInSec);
        var hourLeft = 0, minLeft = 0, ssLeft = 0;
        if (isCountUp) {
            var timePassed = currTimeInSec - startTimeInSec;
            hourLeft = Math.floor(timePassed / 3600);
            minLeft = Math.floor((timePassed % 3600) / 60);
            ssLeft = timePassed - hourLeft * 3600 - minLeft * 60;
        } else {
            if (timeLeft == 0) {
                if (timeup_audio != null) {
                    timeup_audio.play();
                }
            } else if (timeLeft > 0) {
                hourLeft = Math.floor(timeLeft / 3600);
                minLeft = Math.floor((timeLeft % 3600) / 60);
                ssLeft = timeLeft - hourLeft * 3600 - minLeft * 60;
            } else {
                isTimeUp = true;
                isTimerStarted = false;
            }
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

    function drawBarClock(posY) {
        var rectw = canvasWidth * 0.8;
        var recth = canvasHeight / 5;
        drawBarClockBody(rectw, recth, posY);
        drawBarClockNumber(rectw, recth);
        drawBarClockPointer(rectw, recth);
    }

    function drawBarClockBody(rectw, recth, posY) {
        barPtrx = canvasWidth / 2 - rectw / 2;
        barPtry = posY;
        ctx.beginPath();
        ctx.rect(canvasWidth / 2 - rectw / 2, posY, rectw, recth);
        ctx.fillStyle = "#104B10";
        ctx.fill();
    }
    function drawBarClockPointer(rectw, recth) {
        if (isTimerStarted) {
            var totalDuration = HH * 3600 + MM * 60 + SS;
            var startTimeInSec = startTime == null ? 0 : startTime.getHours() * 3600 + startTime.getMinutes() * 60 + startTime.getSeconds() + pausedDuration;
            var now = new Date();
            var currTimeInSec = now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds();
            if (currTimeInSec - startTimeInSec == totalDuration) {
                if (timeup_audio != null) {
                    timeup_audio.play();
                }
            }
            if (currTimeInSec - startTimeInSec >= totalDuration) {
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


    // #region DigiBarClock
    function drawDigiBarClock(posY1, posY2) {
        drawDigitalClock(posY1);
        drawBarClock(posY2);
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
                if (showCombi) {
                    drawDigiBarClock(canvasHeight / 4, canvasHeight / 2);
                } else {
                    drawBarClock(3 * canvasHeight / 10);
                }
                break;
            case ClockType.DIGITAL_CLOCK:
                if (showCombi) {
                    drawDigiBarClock(canvasHeight / 4, canvasHeight / 2);
                } else {
                    drawDigitalClock(canvasHeight / 2);
                }
                break;
            case ClockType.SQUARE_CLOCK:
                var size = Math.min.apply(null, [canvasHeight, canvasWidth]) / 2;
                radius = size * 0.9;
                ctx.translate(size, size);
                drawSquareClock();
                ctx.translate(-size, -size);
                break;
        }
        if (isTimerStarted && tick_audio != null) {
            tick_audio.play();
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
        isTimerStarted = false;
        isTimeUp = false;
        if (showCombi) {
            $("#checkbox-digital-bar-clock").trigger('click');
        }

        if (isCountUp && clockType == ClockType.DIGITAL_CLOCK) {
            $("#checkbox-digital-count-up").trigger('click');
        }
        $("#toolbar-HH").val(HH).trigger("change");
        $("#toolbar-MM").val(MM).trigger("change");
        $("#toolbar-SS").val(SS).trigger("change");
        $('#toolbar-interval').val(interval).trigger("change");
        saveSettings('HH', HH);
        saveSettings('MM', MM);
        saveSettings('SS', SS);
        saveSettings('interval', interval);
        saveSettings('showCombi', false);
        saveSettings('isCountUp', false);
        showCombi = false;
        isCountUp = false;
        isReset = false;
        pausedDuration = 0;
        pausedTimeInSec = 0;
        isPaused = false;
        clearInterval(timer);
        clearInterval(pausedTimer);
    }

    function loadTickSound(t) {
        switch (tickType) {
            case TickType.NONE:
                tick_audio = null;
                break;
            case TickType.TICK:
                tick_audio = new Audio('../Resources/Audio/ticking.mp3');
                break;
        }
    }

    function loadTimeupSound() {
        switch (timeupType) {
            case TimeupType.NONE:
                timeup_audio = null;
                break;
            case TimeupType.ALARM:
                timeup_audio = new Audio('../Resources/Audio/alarm.mp3');
                break;
        }
    }

    function updatePauseDuration() {
        pausedDuration += 1;
    }

    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    // #endregion
})();