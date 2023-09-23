(function($){
    $(document).ready(function(){
        function formatNumber(input){
            input.value = parseFloat(input.value).toLocaleString('en-US');
        }

        // Target input field by id or class
        const decimalFields = document.querySelectorAll('.decimal-field');

        decimalFields.forEach(function(field){
            field.addEventListener('focusout', function(event){
                formatNumber(event.target);
            });
        });
    });
})(django.jQuery);
