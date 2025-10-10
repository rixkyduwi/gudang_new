!(function ($) {
    "use strict";
  
    var CalendarApp = function () {
      this.$body = $("body");
      (this.$calendar = $("#calendar")),
        (this.$event = "#calendar-events div.calendar-events"),
        (this.$categoryForm = $("#add-new-event form")),
        (this.$extEvents = $("#calendar-events")),
        (this.$modal = $("#my-event")),
        (this.$saveCategoryBtn = $(".save-category")),
        (this.$calendarObj = null);
    };
    function init_date(dateStringWithTime) {
      // Pisahkan tanggal, jam, dan menit dari string
      var dateTimeParts = dateStringWithTime.split(", ");
      var datePart = dateTimeParts[0];
      var timePart = dateTimeParts[1];
  
      var dateParts = datePart.split("/");
      var day = parseInt(dateParts[0], 10);
      var month = parseInt(dateParts[1], 10) - 1; // Bulan dimulai dari 0 (Januari = 0, Februari = 1, ...)
      var year = parseInt(dateParts[2], 10);
  
      var timeParts = timePart.split(".");
      var hours = parseInt(timeParts[0], 10);
      var minutes = parseInt(timeParts[1], 10);
  
      // Membuat objek Date dengan tanggal, jam, dan menit yang ditentukan
      var dateObject = new Date(year, month, day, hours, minutes);
  
      return dateObject;
    }
    function reverse_datetime_local(dateStringWithTime) {
      // Membuat objek Date dari string input
      var date = new Date(dateStringWithTime);
  
      // Mendapatkan komponen waktu
      var day = date.getDate();
      var month = date.getMonth() + 1; // bulan dimulai dari 1
      var year = date.getFullYear();
      var hours = date.getHours()-1;
      var minutes = date.getMinutes()-1;
      hours = hours.toString().padStart(2, "0");
      minutes = minutes.toString().padStart(2, "0");
  
      // Mengatur format sesuai keinginan
      var formattedDate = `${day}/${month}/${year}, ${hours}:${minutes}`;
  
      return formattedDate;
    }
    function init_datetime_local(dateStringWithTime) {
      // Pisahkan tanggal, jam, dan menit dari string
      var dateTimeParts = dateStringWithTime.split(", ");
      var datePart = dateTimeParts[0];
      var timePart = dateTimeParts[1];
  
      var dateParts = datePart.split("/");
      var day = parseInt(dateParts[0], 10).toString().padStart(2, "0");
      var month = parseInt(dateParts[1], 10).toString().padStart(2, "0"); // Bulan dimulai dari 1 (Januari = 1, Februari = 2, ...)
      var year = parseInt(dateParts[2], 10);
  
      var timeParts = timePart.split(".");
      var hours = (parseInt(timeParts[0], 10) + 1).toString().padStart(2, "0");
      var minutes = (parseInt(timeParts[1], 10) + 1).toString().padStart(2, "0");
  
      // Membuat objek Date dengan tanggal, jam, dan menit yang ditentukan
      var dateObject =
        year + "-" + month + "-" + day + "T" + hours + ":" + minutes;
  
      return dateObject;
    }
      /* on click on event */
      (CalendarApp.prototype.onEventClick = function (calEvent, jsEvent, view) {
        // edit dan hapus agenda
        var $this = this;
        $this.$modal.find(".modal-title").empty().append("Agenda");
        var form = $(`<div style="margin: 0px">
            <h4 class="card-title text-center"> `+calEvent.title+`</h4>
            `);
      if (calEvent.gambar == "default.jpg"){}
               else {
              form.append(`<img class="card-img img-responsive" src="../../static/image/`+calEvent.gambar+`" alt="Card image cap" />`)
               }
               form.append(`Kategori : `+ calEvent.kategori+`<br>
               Dari : `+ reverse_datetime_local(calEvent.jam_mulai) +`<br>
               Sampai : `+ reverse_datetime_local(calEvent.jam_selesai) +`<br>
               Pelaksana : `+ calEvent.pemimpin_kegiatan +`<br>
               `+ calEvent.keterangan +`
           
         </div>`)
       
          $this.$modal.show();
          $(".bckdrop").addClass("show");
          $(".bckdrop").removeClass("hide");
          $this.$modal
            .find(".modal-body")
            .empty()
            .prepend(form)
            .end()
          $this.$modal.find(".close-dialog").click(function () {
              $this.$modal.hide("hide");
              $(".bckdrop").addClass("hide");
              $(".bckdrop").removeClass("show");
              $("body").removeClass("modal-open");
            });
            $("body").addClass("modal-open");
           
            $this.$calendarObj.fullCalendar("unselect");
        }),
      /* on select */
      (CalendarApp.prototype.onSelect = function (start, end, allDay) {
        // tambah agenda
        var $this = this;
        $this.$modal.show();
        $this.$modal.find(".modal-title").empty().append("Agenda");
        $(".bckdrop").addClass("show");
        $(".bckdrop").removeClass("hide");
        var msg = $("<p> Belum Ada Agenda </p>");
        
        $this.$modal
        .find(".modal-body")
        .empty()
        .prepend(msg)
        .end()
      $this.$modal.find(".close-dialog").click(function () {
        $this.$modal.hide("hide");
        $(".bckdrop").addClass("hide");
        $(".bckdrop").removeClass("show");
        $("body").removeClass("modal-open");
      });
      $("body").addClass("modal-open");
     
      $this.$calendarObj.fullCalendar("unselect");

      }),
      (CalendarApp.prototype.enableDrag = function () {
        //init events
        $(this.$event).each(function () {
          // create an Event Object (http://arshaw.com/fullcalendar/docs/event_data/Event_Object/)
          // it doesn't need to have a start or end
          var eventObject = {
            title: $.trim($(this).text()), // use the element's text as the event title
          };
          // store the Event Object in the DOM element so we can get to it later
          $(this).data("eventObject", eventObject);
          // make the event draggable using jQuery UI
          $(this).draggable({
            zIndex: 999,
            revert: true, // will cause the event to go back to its
            revertDuration: 0, //  original position after the drag
          });
        });
      });
  
    /* Initializing */
    (CalendarApp.prototype.init = function () {
      // ambil db
      this.enableDrag();
      /*  Initialize the calendar  */
      var date = new Date();
      var d = date.getDate();
      var m = date.getMonth();
      var y = date.getFullYear();
      var form = "";
      var today = new Date($.now());
  
      var DataDB = document
        .getElementById("data-container")
        .getAttribute("data-db");
      DataDB = DataDB.replace(/'/g, '"');
      DataDB = DataDB.replace("'", '"');
      console.log(DataDB);
      var defaultEvents = [];
      var jsonData = JSON.parse(DataDB);
      for (var i = 0; i < jsonData.length; i++) {
        // Mengambil nilai id dan total dari setiap objek
        var title = jsonData[i].title;
        console.log(title);
        var start = jsonData[i].start;
        var end = jsonData[i].end;
        var kategori = jsonData[i].kategori;
        var className = "";
        if (kategori == "darurat") {
          className = "bg-light-danger border-start border-2 border-danger";
        } else if (kategori == "penting") {
          className = "bg-warning border-start border-2 border-warning";
        } else if (kategori == "event") {
          className = "bg-light-success border-start border-2 border-success";
        } else if (kategori == "pemberitahuan") {
          className = "bg-light-warning border-start border-2 border-warning";
        } else if (kategori == "pemilihan") {
          className = "bg-light-primary border-start border-2 border-primary";
        }
  
        let agenda = {
          title: title,
          start: init_date(start),
          end: init_date(end),
          jam_mulai: init_datetime_local(start),
          jam_selesai: init_datetime_local(end),
          gambar:jsonData[i].gambar,
          className: className,
          id: jsonData[i].id,
          keterangan: jsonData[i].keterangan,
          foto: jsonData[i].foto,
          kategori: jsonData[i].kategori,
          pemimpin_kegiatan: jsonData[i].pemimpin_kegiatan,
        };
        defaultEvents.push(agenda);
      }
      var $this = this;
      $this.$calendarObj = $this.$calendar.fullCalendar({
        slotDuration: "00:15:00",
        /* If we want to split day time each 15minutes */
        minTime: "08:00:00",
        maxTime: "19:00:00",
        currentText: "hari ini",
        buttonText: {
          month: "Bulan",
          week: "Minggu",
          day: "Hari",
          list: "Agenda",
        },
        defaultView: "month",
        handleWindowResize: true,
  
        header: {
          left: "prev,next today",
          center: "title",
          right: "month,agendaWeek",
        },
        events: defaultEvents,
        editable: true,
        droppable: false, // this allows things to be dropped onto the calendar !!!
        selectable: true,
        drop: function (date) {
          // hapus event agenda
          $this.onDrop($(this), date);
        },
        select: function (start, end, allDay) {
          // edit event agenda
          $this.onSelect(start, end, allDay);
        },
        eventClick: function (calEvent, jsEvent, view) {
          //tambah event agenda
          $this.onEventClick(calEvent, jsEvent, view);
        },
      });
  
      //on new event
      this.$saveCategoryBtn.on("click", function () {
        var categoryName = $this.$categoryForm
          .find("input[name='category-name']")
          .val();
        var categoryColor = $this.$categoryForm
          .find("select[name='category-color']")
          .val();
        if (categoryName !== null && categoryName.length != 0) {
          $this.$extEvents.append(
            '<div class="calendar-events mb-3" data-class="bg-' +
              categoryColor +
              '" style="position: relative;"><i class="fa fa-circle text-' +
              categoryColor +
              ' me-2" ></i>' +
              categoryName +
              "</div>"
          );
          $this.enableDrag();
        }
      });
    }),
      //init CalendarApp
      ($.CalendarApp = new CalendarApp()),
      ($.CalendarApp.Constructor = CalendarApp);
  })(window.jQuery),
    //initializing CalendarApp
    $(window).on("load", function () {
      $.CalendarApp.init();
      //fix today button
      $(".fc-today-button").html("hari ini");
      mobilesize("load");
    });
    function mobilesize(status){
      var moreEventsLink = $('.fc-day-grid-event');
      var windowWidth = $(window).width();
      if (windowWidth < 480) {
        moreEventsLink.removeClass('fc-h-event')
        moreEventsLink.removeClass('fc-event')
        moreEventsLink.removeClass('fc-start')
        moreEventsLink.removeClass('fc-end')
        moreEventsLink.removeClass('bg-light-danger')
        moreEventsLink.removeClass('border-start')
        moreEventsLink.removeClass('border-2')
        moreEventsLink.removeClass('border-danger')
        moreEventsLink.removeClass('fc-draggable');
        moreEventsLink.html("+1 agenda"); // Change text to "+1 more"
      }
    }
    $(window).on('resize', function() {
     mobilesize("resize")
    });

    // Trigger resize event to set initial state
    $(window).resize();
  