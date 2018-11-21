const api = require('./api');


module.exports = async function (activity) {

    try {

        api.initialize(activity);  
 
        // *todo* create a simple and fast API request
        const response = await api('/me');
      
        // ping success: when result returns status 200
        activity.Response.Data = { success: (response && response.statusCode == 200) };


    } catch (error) {

        // in ping errors are returned in regular response with success == false
        var response = { success:  false};

        // return error response
        var m = error.message;    
        if (error.stack) m = m + ": " + error.stack;

        response.ErrorText = m;
        activity.Response.Data = response;
    }

};
