jQuery( document ).ready(function() {
  const picker = MCDatepicker.create({
	  el: '#date_start'
  })

  jQuery("#create_event").validate({
    rules: {
      'language[]': {
        require_from_group: [1, ".language_group"]
      }
    }
  });
});
