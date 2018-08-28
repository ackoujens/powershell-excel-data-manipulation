class SQLObject {
    [string] $server;
    [string] $database;
    [string] $sqlCommand;
    [string] $connectionString;
    $sqlConnection;

    # Initialize object based on filepath
    SQLObject([string] $server, [string] $database) {
        # $this.cleanup();
        $this.server = $server;
        $this.database = $database;
        $this.connect();
    }

    [void] connect() {
        $this.sqlConnection = New-Object System.Data.SqlClient.SqlConnection;
        $this.sqlConnection.ConnectionString = "Server=" + $this.server + ";Database=" + $this.database + ";Integrated Security = true;"
    }

    [void] setQuery([string] $query) {
        $this.sqlCommand = $query;
    }

    [object] execQuery() {
        $connection = new-object system.data.SqlClient.SQLConnection($this.sqlConnection.connectionString);
        $command = new-object system.data.sqlclient.sqlcommand($this.sqlCommand, $connection);
        $connection.Open();

        $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $command;
        $dataset = New-Object System.Data.DataSet;
        $adapter.Fill($dataSet) | Out-Null;

        #$dataSet.Tables;
        $result = $DataSet.Tables
        $connection.Close();

        #return $list;
        return $result

        $this.setQuery("");
    }


    # TODO: Close all open SQL connections/processes if exists
    # [void] cleanup() {}
}