/* Check form fields for UV Index lookup by ZIP Code or city and state (Apr 2009) */

function validateInput() {
	var f = document.forms["uviform"];
	if ((typeof f == "undefined")
	||  (typeof f.zipcode == "undefined")
	||  (typeof f.city_name == "undefined")
	||  (typeof f.state_code == "undefined")) {
		return true;  /* fields not present to check, allow blind form submission */
	}
	if (f.zipcode.value.length == 0 && f.city_name.value.length == 0 && f.state_code.value.length == 0) {
		alert("Enter either a ZIP Code or a city and state.");
		f.zipcode.focus();
		return false;
	}
	if (f.zipcode.value.length > 0) {
		var zipre = /^[0-9]+$/;
		if ((f.zipcode.value.length != 5) || (!zipre.test(f.zipcode.value))) {
			alert("ZIP Code must be 5 digits.");
			f.zipcode.focus();
			return false;
		}
		else {
			return true;
		}
	}
	else {
		var stre = /^[a-zA-Z]+$/;
		if ((f.city_name.value.length == 0) || (f.state_code.value.length == 0)) {
			alert("Enter both city and state.");
			f.city_name.focus();
			return false;
		}
		else if ((f.state_code.value.length != 2) || (!stre.test(f.state_code.value))) {
			alert("State must be a 2-letter postal code");
			f.state_code.focus();
			return false;
		}
		else {
			return true;
		}
	}
}
