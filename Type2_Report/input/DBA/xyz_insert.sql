-- Note that this is a sample PL/SQL script for DBA tasks. 
-- In the actual work environment, there are hundreds of similar SQL files in the input filder automated by the program.
-- All the actual database and column names has been replaced with XXXXX.

DECLARE
    -- Data warehouse variable, inserted during Python automation
    data_warehouse VARCHAR2(10) := DATA_WAREHOUSE_VALUE;
BEGIN
    -- Insert into XXXXX_subsystem table for PRD and TST environments
    INSERT INTO XXX_XXX_PRD.XXXXX_subsystem (XXXXX_CLIENT_OID, SCHEMA_NAME) 
    VALUES (XXXXX_CLIENT_OID_VALUE, SCHEMA_NAME_VALUE_PRD);

    INSERT INTO XXX_XXX_TST.XXXXX_subsystem (XXXXX_CLIENT_OID, SCHEMA_NAME) 
    VALUES (XXXXX_CLIENT_OID_VALUE, SCHEMA_NAME_VALUE_TST);

    -- Insert into XXXXX_MEASURES table
    INSERT INTO XXX_XXX_PRD.XXXXX_MEASURES
    (XXXXX_CLIENT_OID, DIST_ID, XXXXX_OID, XXXXX_TYPE)
    VALUES (XXXXX_CLIENT_OID_VALUE, 0, 3, 'FXXXX_XXXXXX');

    -- Insert into XXXXX_XXXXX_XXXXX_SCHEMA table
    INSERT INTO XXX_XXX_PRD.XXXXX_XXXXX_XXXXX_SCHEMA
    (XXXXX_CLIENT_OID, WAREHOUSE_NUMBER, SCHEMA_NAME)
    VALUES (XXXXX_CLIENT_OID_VALUE, 0, 'XYZ_SUB_ODS_PRD');

    -- DB links must be created before executing the following insert statements
    IF data_warehouse = '1' THEN
        -- If data warehouse is 1, insert into XXXXX_XXXXX_XXXXX_SCHEMA with specific schema name
        INSERT INTO XXX_XXX_PRD.XXXXX_XXXXX_XXXXX_SCHEMA
        (XXXXX_CLIENT_OID,WAREHOUSE_NUMBER,SCHEMA_NAME)
        VALUES (XXXXX_CLIENT_OID_VALUE, 3, '@TO_XYZ_SUB_DWH_PRD.XXXXX.COM');

    ELSIF data_warehouse = '2' THEN
        -- If data warehouse is 2, insert into XXXXX_XXXXX_XXXXX_SCHEMA with specific schema name
        INSERT INTO XXX_XXX_PRD.XXXXX_XXXXX_XXXXX_SCHEMA
        (XXXXX_CLIENT_OID,WAREHOUSE_NUMBER,SCHEMA_NAME)
        VALUES (XXXXX_CLIENT_OID_VALUE, 4, '@TO_XYZ_SUB_DWH_PR2.XXXXX.COM');

    ELSIF data_warehouse = 'BOTH' THEN
        -- If data warehouse is BOTH, insert into XXXXX_XXXXX_XXXXX_SCHEMA with specific schema names
        INSERT INTO XXX_XXX_PRD.XXXXX_XXXXX_XXXXX_SCHEMA
        (XXXXX_CLIENT_OID,WAREHOUSE_NUMBER,SCHEMA_NAME)
        VALUES (XXXXX_CLIENT_OID_VALUE, 3, '@TO_XYZ_SUB_DWH_PRD.XXXXX.COM');

        INSERT INTO XXX_XXX_PRD.XXXXX_XXXXX_XXXXX_SCHEMA
        (XXXXX_CLIENT_OID,WAREHOUSE_NUMBER,SCHEMA_NAME)
        VALUES (XXXXX_CLIENT_OID_VALUE, 4, '@TO_XYZ_SUB_DWH_PR2.XXXXX.COM');

    END IF;

    COMMIT;
END;
/
