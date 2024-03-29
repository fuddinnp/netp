$(document).ready(function () {
  $(document).on("opening", ".remodal", function () {
    setTimeout(function () {
      $("html").addClass("ios-lock-fix");
    }, 300);
  });
});

$(document).ready(function () {
  $(document).on("closing", ".remodal", function () {
    setTimeout(function () {
      $("html").removeClass("ios-lock-fix");
    }, 300);
  });
});

var btn = $("#button");

$(window).scroll(function () {
  if ($(window).scrollTop() > 500) {
    btn.addClass("show");
  } else {
    btn.removeClass("show");
  }
});

btn.on("click", function (e) {
  e.preventDefault();
  $("html, body").animate({ scrollTop: 0 }, "300");
});

// Get the modal
$(document).ready(function () {
  var modal = document.getElementById("myModal");

  // Get the button that opens the modal
  var btn = document.getElementById("myBtn");

  // Get the <span> element that closes the modal
  var span = document.getElementsByClassName("close")[0];

  // When the user clicks the button, open the modal
  btn.onclick = function () {
    modal.style.display = "block";
  };

  // When the user clicks on <span> (x), close the modal
  span.onclick = function () {
    modal.style.display = "none";
  };

  // When the user clicks anywhere outside of the modal, close it
  window.onclick = function (event) {
    if (event.target == modal) {
      modal.style.display = "none";
    }
  };
});

var imgLiquid = imgLiquid || { VER: "0.9.944" };
imgLiquid.bgs_Available = false;
imgLiquid.bgs_CheckRunned = false;
imgLiquid.injectCss = ".imgLiquid img {visibility:hidden}";

(function ($) {
  // ___________________________________________________________________

  function checkBgsIsavailable() {
    if (imgLiquid.bgs_CheckRunned) return;
    else imgLiquid.bgs_CheckRunned = true;

    var spanBgs = $('<span style="background-size:cover" />');
    $("body").append(spanBgs);

    !(function () {
      var bgs_Check = spanBgs[0];
      if (!bgs_Check || !window.getComputedStyle) return;
      var compStyle = window.getComputedStyle(bgs_Check, null);
      if (!compStyle || !compStyle.backgroundSize) return;
      imgLiquid.bgs_Available = compStyle.backgroundSize === "cover";
    })();

    spanBgs.remove();
  }

  // ___________________________________________________________________

  $.fn.extend({
    imgLiquid: function (options) {
      this.defaults = {
        fill: true,
        verticalAlign: "center", //	'top'	//	'bottom' // '50%'  // '10%'
        horizontalAlign: "center", //	'left'	//	'right'  // '50%'  // '10%'
        useBackgroundSize: true,
        useDataHtmlAttr: true,

        responsive: true /* Only for use with BackgroundSize false (or old browsers) */,
        delay: 0 /* Only for use with BackgroundSize false (or old browsers) */,
        fadeInTime: 0 /* Only for use with BackgroundSize false (or old browsers) */,
        removeBoxBackground: true /* Only for use with BackgroundSize false (or old browsers) */,
        hardPixels: true /* Only for use with BackgroundSize false (or old browsers) */,
        responsiveCheckTime: 500 /* Only for use with BackgroundSize false (or old browsers) */ /* time to check div resize */,
        timecheckvisibility: 500 /* Only for use with BackgroundSize false (or old browsers) */ /* time to recheck if visible/loaded */,

        // CALLBACKS
        onStart: null, // no-params
        onFinish: null, // no-params
        onItemStart: null, // params: (index, container, img )
        onItemFinish: null, // params: (index, container, img )
        onItemError: null, // params: (index, container, img )
      };

      checkBgsIsavailable();
      var imgLiquidRoot = this;

      // Extend global settings
      this.options = options;
      this.settings = $.extend({}, this.defaults, this.options);

      // CallBack
      if (this.settings.onStart) this.settings.onStart();

      // ___________________________________________________________________

      return this.each(function ($i) {
        // MAIN >> each for image

        var settings = imgLiquidRoot.settings,
          $imgBoxCont = $(this),
          $img = $("img:first", $imgBoxCont);
        if (!$img.length) {
          onError();
          return;
        }

        // Extend settings
        if (!$img.data("imgLiquid_settings")) {
          // First time
          settings = $.extend(
            {},
            imgLiquidRoot.settings,
            getSettingsOverwrite()
          );
        } else {
          // Recall
          // Remove Classes
          $imgBoxCont
            .removeClass("imgLiquid_error")
            .removeClass("imgLiquid_ready");
          settings = $.extend(
            {},
            $img.data("imgLiquid_settings"),
            imgLiquidRoot.options
          );
        }
        $img.data("imgLiquid_settings", settings);

        // Start CallBack
        if (settings.onItemStart)
          settings.onItemStart($i, $imgBoxCont, $img); /* << CallBack */

        // Process
        if (imgLiquid.bgs_Available && settings.useBackgroundSize)
          processBgSize();
        else processOldMethod();

        // END MAIN <<

        // ___________________________________________________________________

        function processBgSize() {
          // Check change img src
          if (
            $imgBoxCont
              .css("background-image")
              .indexOf(encodeURI($img.attr("src"))) === -1
          ) {
            // Change
            $imgBoxCont.css({
              "background-image": 'url("' + encodeURI($img.attr("src")) + '")',
            });
          }

          $imgBoxCont.css({
            "background-size": settings.fill ? "cover" : "contain",
            "background-position": (
              settings.horizontalAlign +
              " " +
              settings.verticalAlign
            ).toLowerCase(),
            "background-repeat": "no-repeat",
          });

          $("a:first", $imgBoxCont).css({
            display: "block",
            width: "100%",
            height: "100%",
          });

          $("img", $imgBoxCont).css({ display: "none" });

          if (settings.onItemFinish)
            settings.onItemFinish($i, $imgBoxCont, $img); /* << CallBack */

          $imgBoxCont.addClass("imgLiquid_bgSize");
          $imgBoxCont.addClass("imgLiquid_ready");
          checkFinish();
        }

        // ___________________________________________________________________

        function processOldMethod() {
          // Check change img src
          if ($img.data("oldSrc") && $img.data("oldSrc") !== $img.attr("src")) {
            /* Clone & Reset img */
            var $imgCopy = $img.clone().removeAttr("style");
            $imgCopy.data(
              "imgLiquid_settings",
              $img.data("imgLiquid_settings")
            );
            $img.parent().prepend($imgCopy);
            $img.remove();
            $img = $imgCopy;
            $img[0].width = 0;

            // Bug ie with > if (!$img[0].complete && $img[0].width) onError();
            setTimeout(processOldMethod, 10);
            return;
          }

          // Reproceess?
          if ($img.data("imgLiquid_oldProcessed")) {
            makeOldProcess();
            return;
          }

          // Set data
          $img.data("imgLiquid_oldProcessed", false);
          $img.data("oldSrc", $img.attr("src"));

          // Hide others images
          $("img:not(:first)", $imgBoxCont).css("display", "none");

          // CSSs
          $imgBoxCont.css({ overflow: "hidden" });
          $img.fadeTo(0, 0).removeAttr("width").removeAttr("height").css({
            visibility: "visible",
            "max-width": "none",
            "max-height": "none",
            width: "auto",
            height: "auto",
            display: "block",
          });

          // CheckErrors
          $img.on("error", onError);
          $img[0].onerror = onError;

          // loop until load
          function onLoad() {
            if (
              $img.data("imgLiquid_error") ||
              $img.data("imgLiquid_loaded") ||
              $img.data("imgLiquid_oldProcessed")
            )
              return;
            if (
              $imgBoxCont.is(":visible") &&
              $img[0].complete &&
              $img[0].width > 0 &&
              $img[0].height > 0
            ) {
              $img.data("imgLiquid_loaded", true);
              setTimeout(makeOldProcess, $i * settings.delay);
            } else {
              setTimeout(onLoad, settings.timecheckvisibility);
            }
          }

          onLoad();
          checkResponsive();
        }

        // ___________________________________________________________________

        function checkResponsive() {
          /* Only for oldProcessed method (background-size dont need) */

          if (!settings.responsive && !$img.data("imgLiquid_oldProcessed"))
            return;
          if (!$img.data("imgLiquid_settings")) return;

          settings = $img.data("imgLiquid_settings");

          $imgBoxCont.actualSize =
            $imgBoxCont.get(0).offsetWidth +
            $imgBoxCont.get(0).offsetHeight / 10000;
          if (
            $imgBoxCont.sizeOld &&
            $imgBoxCont.actualSize !== $imgBoxCont.sizeOld
          )
            makeOldProcess();

          $imgBoxCont.sizeOld = $imgBoxCont.actualSize;
          setTimeout(checkResponsive, settings.responsiveCheckTime);
        }

        // ___________________________________________________________________

        function onError() {
          $img.data("imgLiquid_error", true);
          $imgBoxCont.addClass("imgLiquid_error");
          if (settings.onItemError)
            settings.onItemError($i, $imgBoxCont, $img); /* << CallBack */
          checkFinish();
        }

        // ___________________________________________________________________

        function getSettingsOverwrite() {
          var SettingsOverwrite = {};

          if (imgLiquidRoot.settings.useDataHtmlAttr) {
            var dif = $imgBoxCont.attr("data-imgLiquid-fill"),
              ha = $imgBoxCont.attr("data-imgLiquid-horizontalAlign"),
              va = $imgBoxCont.attr("data-imgLiquid-verticalAlign");

            if (dif === "true" || dif === "false")
              SettingsOverwrite.fill = Boolean(dif === "true");
            if (
              ha !== undefined &&
              (ha === "left" ||
                ha === "center" ||
                ha === "right" ||
                ha.indexOf("%") !== -1)
            )
              SettingsOverwrite.horizontalAlign = ha;
            if (
              va !== undefined &&
              (va === "top" ||
                va === "bottom" ||
                va === "center" ||
                va.indexOf("%") !== -1)
            )
              SettingsOverwrite.verticalAlign = va;
          }

          if (imgLiquid.isIE && imgLiquidRoot.settings.ieFadeInDisabled)
            SettingsOverwrite.fadeInTime = 0; //ie no anims
          return SettingsOverwrite;
        }

        // ___________________________________________________________________

        function makeOldProcess() {
          /* Only for old browsers, or useBackgroundSize seted false */
          // Calculate size
          var w,
            h,
            wn,
            hn,
            ha,
            va,
            hdif,
            vdif,
            margT = 0,
            margL = 0,
            $imgCW = $imgBoxCont.width(),
            $imgCH = $imgBoxCont.height();

          // Save original sizes
          if ($img.data("owidth") === undefined)
            $img.data("owidth", $img[0].width);
          if ($img.data("oheight") === undefined)
            $img.data("oheight", $img[0].height);

          // Compare ratio
          if (
            settings.fill ===
            $imgCW / $imgCH >= $img.data("owidth") / $img.data("oheight")
          ) {
            w = "100%";
            h = "auto";
            wn = Math.floor($imgCW);
            hn = Math.floor(
              $imgCW * ($img.data("oheight") / $img.data("owidth"))
            );
          } else {
            w = "auto";
            h = "100%";
            wn = Math.floor(
              $imgCH * ($img.data("owidth") / $img.data("oheight"))
            );
            hn = Math.floor($imgCH);
          }

          // Align X
          ha = settings.horizontalAlign.toLowerCase();
          hdif = $imgCW - wn;
          if (ha === "left") margL = 0;
          if (ha === "center") margL = hdif * 0.5;
          if (ha === "right") margL = hdif;
          if (ha.indexOf("%") !== -1) {
            ha = parseInt(ha.replace("%", ""), 10);
            if (ha > 0) margL = hdif * ha * 0.01;
          }

          // Align Y
          va = settings.verticalAlign.toLowerCase();
          vdif = $imgCH - hn;
          if (va === "top") margT = 0;
          if (va === "center") margT = vdif * 0.5;
          if (va === "bottom") margT = vdif;
          if (va.indexOf("%") !== -1) {
            va = parseInt(va.replace("%", ""), 10);
            if (va > 0) margT = vdif * va * 0.01;
          }

          // Add Css
          if (settings.hardPixels) {
            w = wn;
            h = hn;
          }
          $img.css({
            width: w,
            height: h,
            "margin-left": Math.floor(margL),
            "margin-top": Math.floor(margT),
          });

          // FadeIn > Only first time
          if (!$img.data("imgLiquid_oldProcessed")) {
            $img.fadeTo(settings.fadeInTime, 1);
            $img.data("imgLiquid_oldProcessed", true);
            if (settings.removeBoxBackground)
              $imgBoxCont.css("background-image", "none");
            $imgBoxCont.addClass("imgLiquid_nobgSize");
            $imgBoxCont.addClass("imgLiquid_ready");
          }

          if (settings.onItemFinish)
            settings.onItemFinish($i, $imgBoxCont, $img); /* << CallBack */
          checkFinish();
        }

        // ___________________________________________________________________

        function checkFinish() {
          /* Check callBack */ if ($i === imgLiquidRoot.length - 1)
            if (imgLiquidRoot.settings.onFinish)
              imgLiquidRoot.settings.onFinish();
        }
      });
    },
  });
})(jQuery);

// Inject css styles ______________________________________________________
!(function () {
  var css = imgLiquid.injectCss,
    head = document.getElementsByTagName("head")[0],
    style = document.createElement("style");
  style.type = "text/css";
  if (style.styleSheet) {
    style.styleSheet.cssText = css;
  } else {
    style.appendChild(document.createTextNode(css));
  }
  head.appendChild(style);
})();

/********** Custom Scrollbar ***************/

(function ($) {
  $(window).on("load", function () {
    $(".scrollbar").mCustomScrollbar();
  });
})(jQuery);

/********** Custom Scrollbar ***************/

$(document).ready(function () {
  $("#landing-service").owlCarousel({
    stagePadding: 50,
    margin: 0,
    center: true,
    loop: true,
    responsiveClass: true,
    dots: false,
    nav: true,

    responsive: {
      0: {
        items: 1,
        nav: true,
        stagePadding: 20,
      },
      768: {
        items: 3,
        nav: false,
      },
    },
  });
});
$(document).ready(function () {
  $(".case-study-safeway, .case-study-verizon, .case-study-cisco").owlCarousel({
    loop: false,
    margin: 10,
    responsiveClass: true,
    items: 2,

    mouseDrag: false,
    responsive: {
      0: {
        items: 1,
        nav: true,
      },
      768: {
        items: 2,
        nav: false,
      },

      1024: {
        items: 2, // from 768 screen width to 1024 8 items
        slideBy: 2,
      },
    },
  });
});

$(document).ready(function () {
  $(".imgLiquidFill").imgLiquid();
});

$(document).ready(function () {
  $(".owl-next,.owl-prev").click(function () {
    $(".slider-title.verizon")
      .find("a")
      .attr(
        "data-remodal-target",
        $(".case-study-verizon")
          .find(".owl-item.active")
          .find("a")
          .attr("data-remodal-target")
      );
    $(".slider-title.cisco")
      .find("a")
      .attr(
        "data-remodal-target",
        $(".case-study-cisco")
          .find(".owl-item.active")
          .find("a")
          .attr("data-remodal-target")
      );
    $(".slider-title.safeway")
      .find("a")
      .attr(
        "data-remodal-target",
        $(".case-study-safeway")
          .find(".owl-item.active")
          .find("a")
          .attr("data-remodal-target")
      );
  });
});

resetform();

/*************************** CONTACT FORM **********************/
function submitform() {
  document.getElementById("errorMsgemail").style.display = "none";
  document.getElementById("errorMsgcomment").style.display = "none";
  if (validateEmail(document.getElementById("emailAddr").value) == false) {
    document.getElementById("errorMsgemail").style.display = "block";
    return false;
  }

  if (trimIt(document.getElementById("comments").value) == "") {
    document.getElementById("errorMsgcomment").style.display = "block";
    return false;
  }

  document.getElementById("subject").value =
    document.getElementById("emailAddr").value + " from Netpace.com";
  document.myform.submit();
}

function sendEmailAjax() {
  document.getElementById("errorMsgemail").style.display = "none";
  document.getElementById("errorMsgcomment").style.display = "none";
  document.getElementById("successMsg").style.display = "none";
  document.getElementById("errorMsg").style.display = "none";
  document.getElementById("email-span").className = "bar";
  document.getElementById("comment-span").className = "bar";

  if (validateEmail(document.getElementById("emailAddr").value) == false) {
    document.getElementById("errorMsgemail").style.display = "block";
    document.getElementById("email-span").className = "bar-red";
    return false;
  }

  if (trimIt(document.getElementById("comments").value) == "") {
    document.getElementById("errorMsgcomment").style.display = "block";
    document.getElementById("comment-span").className = "bar-red";
    return false;
  }

  document.getElementById("ajax-loader").style.display = "block";

  document.getElementById("subject").value =
    document.getElementById("emailAddr").value + " from Netpace.com";
  var myform = document.getElementById("myform");

  jQuery.ajax({
    url: "formmail/formmail.asp",
    data: jQuery(myform).serialize(),
    type: "POST",
    cache: false,
    success: function (msg) {
      document.getElementById("successMsg").style.display = "block";
      document.getElementById("ajax-loader").style.display = "none";
      document.getElementById("comment-span").className = "bar";
      document.getElementById("email-span").className = "bar";

      document.getElementById("comments").value = "";
      document.getElementById("emailAddr").value = "";
      document.getElementById("name").value = "";
      document.getElementById("pnumber").value = "";
    },
    error: function (xhr, ajaxOptions, thrownError) {
      document.getElementById("errorMsg").style.display = "block";
      document.getElementById("ajax-loader").style.display = "none";
    },
  });
}

function validateEmail(elementValue) {
  if (elementValue == "") return false;

  var emailPattern =
    /^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$/;
  if (emailPattern.test(trimIt(elementValue)) == false) return false;

  return true;
}

function trimIt(elemntV) {
  return elemntV.replace(/^\s\s*/, "").replace(/\s\s*$/, "");
}

// Start Cookies
function setCookie(cname, cvalue, exdays) {
  const cdate = new Date();
  cdate.setTime(cdate.getTime() + exdays * 24 * 60 * 60 * 1000);
  let expires = "expires=" + cdate.toUTCString();
  document.cookie = cname + "=" + cvalue + ";" + expires + ";path=/";
}

function getCookie(cname) {
  let name = cname + "=";
  let decodedCookie = decodeURIComponent(document.cookie);
  let cookieArray = decodedCookie.split(";");
  for (let i = 0; i < cookieArray.length; i++) {
    let cookieObj = cookieArray[i];
    while (cookieObj.charAt(0) == " ") {
      cookieObj = cookieObj.substring(1);
    }
    if (cookieObj.indexOf(name) == 0) {
      return cookieObj.substring(name.length, cookieObj.length);
    }
  }
  return "";
}

$(document).ready(function () {
  try {
    if (getCookieConsentStatus() == null) {
      var cookieDiv = document.getElementById("cookie-div");
      cookieDiv.style.display = "block";
    }
  } catch (e) {
    var cookieDiv = document.getElementById("cookie-div");
    cookieDiv.style.display = "block";
  }
});
function getCookieConsentStatus() {
  return localStorage.getItem("cookies_enabled_netpace");
}

function checkCookie() {
  if (getCookieConsentStatus() == 1) {
    let username = getCookie("username");
    if (username != "") {
      alert("Welcome again " + username);
    } else {
      username = prompt("Please enter your name:", "");
      if (username != "" && username != null) {
        setCookie("username", username, 365);
      }
    }
  }
}

function enableCookies() {
  localStorage.setItem("cookies_enabled_netpace", 1);

  var cookieDiv = document.getElementById("cookie-div");
  cookieDiv.style.display = "none";
}

function disableCookies() {
  localStorage.setItem("cookies_enabled_netpace", 0);

  var cookieDiv = document.getElementById("cookie-div");
  cookieDiv.style.display = "none";
}

// END Cookie
$(document).ready(function () {
  var n = localStorage.getItem("on_load_counter");
  if (n === null) {
    n = 0;
  }
  n = 0;
  // n++;

  localStorage.setItem("on_load_counter", n);
});

function resetform() {
  try {
    document.getElementById("errorMsgemail").style.display = "none";
    document.getElementById("errorMsgcomment").style.display = "none";
    document.getElementById("ajax-loader").style.display = "none";
    document.getElementById("comments").value = "";
    document.getElementById("emailAddr").value = "";
    document.getElementById("name").value = "";
    document.getElementById("pnumber").value = "";
    document.getElementById("successMsg").style.display = "none";
    document.getElementById("errorMsg").style.display = "none";
  } catch (error) {
    console.log(error);
  }

  return false;
}

/*************************************************************************/

(function ($) {
  $(window).on("load", function () {
    $(".content").mCustomScrollbar();
  });
})(jQuery);
