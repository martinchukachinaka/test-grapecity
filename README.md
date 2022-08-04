# test-grapecity

testing grapecity for exporting PDFs and XSLX

## Work Done

1. Added provided functionality
2. included unit tests for services involved
3. Added Integration Tests to on the Similarity and Frequency Controllers

### Before you run

1. Java 11 is required to run this app

2. You need to include the Evaluation license to the application (this is sent along with the email). Two methods to apply:

   a. update the application.yml in src/main/resources/application.yml with the key as shown -

   ```
       grape-city:
           license: ADD-EVALUATION-KEY-HERE
   ```

   b. include the parameter in command line while running the application with the param

   ```
   --args='--grape-city.licence=ADD-EVALUATION-KEY-HERE'
   ```

### How to run the Application

1. run with gradle wrapper

   ```
   ./gradlew bootRun
   ```

   or with this if you've chosen option 2b

   ```
   ./gradlew bootRun --args='--grape-city.licence=ADD-EVALUATION-KEY-HERE'
   ```

2. Application is accessible on port 8686 `http://localhost:8686`.
   This can be changed in the application.yml file

### Note Worthy

1. `/export/xlsx` is used for the xlsx test and current returns a status '200'
2. `/export/pdf` is used for the pdf test and currently fails with a status '500'
3. Both endpoints have their output file at the root of the application or can be changed from the optional parameter &mdash; `outputDirectory`
4. There is an included POSTMAN Collection - `grapecity-tests.postman_collection.json` at the root of the project if that makes it easier to use
