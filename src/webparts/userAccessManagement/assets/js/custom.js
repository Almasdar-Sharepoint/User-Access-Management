function openNav() {
  document.getElementById("mySidenav").style.width = "70%";
  // document.getElementById("flipkart-navbar").style.width = "50%";
  document.body.style.backgroundColor = "rgba(0,0,0,0.4)";
}

function closeNav() {
  document.getElementById("mySidenav").style.width = "0";
  document.body.style.backgroundColor = "rgba(0,0,0,0)";
}
$(document).ready(function () {
  $('.your-class').slick({
      dots: false,
      infinite: true,
      speed: 300,
      slidesToShow: 1,
      centerMode: true,
      variableWidth: true
  });
});


//New Side Bar Navigation and Headar CSS
// Tooltip
var tooltipTriggerList = [].slice.call(document.querySelectorAll('[data-bs-toggle="tooltip"]'))
var tooltipList = tooltipTriggerList.map(function (tooltipTriggerEl) {
  return new bootstrap.Tooltip(tooltipTriggerEl)
})

// Navabar
$('.flipkart-navbar-search .upper-links').click(function(){
  $('li').removeClass("active")
  $(this).addClass("active");
  $(this).next(".border-bottom-warning").show();
});

// Custom toggle Popup
$('.switch').click(function(){
  if($(this).find('.test').prop("checked") == true){
      $(".togglePopup").hide();
  }else{
     $(".togglePopup").show();
  }
});

//Popup close
$('.switch').click(function(){
  if($(this).find('.test').prop("checked") == false){
      $(".togglePopup").hide();
      // $('.test').prop("checked",false);
  }else{
     $(".togglePopup").show();
     $(this).find('.test').addClass('active');
  }
});

//sideNav 
// $(".hamburger-left").on("click", function() {
//   $(".hamburger").toggleClass("is-active")
// })
//Popup deactivate
$('.deactivate').click(function(){
$(".togglePopup").hide();
$('.active').prop("checked",false);
});

$('.close-pop-up').click(function(){
  $(".togglePopup").hide();
});


jQuery(document).on("click",".metismenu .menu-item.sub-menu a",function(){
    
  if(jQuery(this).hasClass("active")){
  jQuery(this).toggleClass("active");
  jQuery(this).closest("li").find(".submenu-dropdown").toggleClass("active");
}
else{
  jQuery(".metismenu .menu-item a.menu-link").removeClass("active");
  jQuery(".metismenu .menu-item .submenu-dropdown").removeClass("active");
  jQuery(this).addClass("active");
  jQuery(this).closest("li").find(".submenu-dropdown").addClass("active");
}
});
