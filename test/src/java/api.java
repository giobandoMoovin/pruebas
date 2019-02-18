
import javax.ws.rs.GET;
import javax.ws.rs.Path;
import javax.ws.rs.core.Response;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

@Path("api")
public class api {

    
    /**
     * Creates a new instance of ApiMoovin
     */
    public api() {
       
    }
    
    @GET
    @Path("/provinces")
    public Response getCostaRicaProvincesArra() {
        int code = 500;
        try {
            code = 200;
        } catch (Exception ex) {
            code = 402;
        }
        return Response.status(code).entity("algo").build();
    }
    
}