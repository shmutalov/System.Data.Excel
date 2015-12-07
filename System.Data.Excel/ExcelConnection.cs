#region Copyright

/*
	Copyright (c) Sherzod Mutalov, 2015
	mailto:shmutalov@gmail.com
*/

#endregion

using System.Data.Common;
using System.Data.Excel.Enums;
using System.Data.Excel.Models;
using System.Data.Excel.Storage;
using System.IO;
using System.Reflection;
using Excel;

namespace System.Data.Excel
{
    /// <summary>
    /// Excel connection
    /// </summary>
    public sealed class ExcelConnection : DbConnection
    {
        #region Properties

        /// <summary>
        /// Used for synchronizing threads
        /// </summary>
        private readonly object _syncLock = new object();

        /// <summary>
        /// Connection string
        /// </summary>
        private string _connectionString;

        /// <summary>
        /// Connection parameters
        /// </summary>
        private ExcelConnectionParameters _parameters;

        /// <summary>
        /// Connection state
        /// </summary>
        private ConnectionState _state;

        /// <summary>
        /// Determines whether object is disposed or not
        /// </summary>
        private bool _disposed;

        /// <summary>
        /// Determines whether object is in disposing state or not
        /// </summary>
        private bool _disposing;

        /// <summary>
        /// Path to storage directory
        /// </summary>
        private string _storegeDir;

        /// <summary>
        /// Internal storage connection
        /// </summary>
        private IDbConnection _storageConnection;

        /// <summary>
        /// Determines whether internal storage initialized or not
        /// </summary>
        private bool _storageInitialized;

        /// <summary>
        /// Storage helper
        /// </summary>
        private readonly IStorage _storage;

        #endregion

        #region Constructors

        public ExcelConnection()
        {
            _state = ConnectionState.Closed;

            _parameters = new ExcelConnectionParameters();

            // initialize default storage
            _storage = new SqliteStorage();

            _storegeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }

        private void InitSettings(string connectionString)
        {
            _parameters = ExcelConnectionParameters.FromConnectionString(connectionString);
            _connectionString = ExcelConnectionParameters.ToConnectionString(_parameters);

            if (string.IsNullOrEmpty(_parameters.StoregeDirectory))
                _storegeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            else
            {
                if (!Directory.Exists(_parameters.StoregeDirectory))
                    throw new ExcelException(
                        "Storage directory '{0}' does not exist",
                        _parameters.StoregeDirectory);

                _storegeDir = _parameters.StoregeDirectory;
            }
        }

        /// <summary>
        /// Creates Excel connection with specified connection string
        /// </summary>
        /// <param name="connectionString"></param>
        public ExcelConnection(string connectionString)
            : this()
        {
            InitSettings(connectionString);
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Creates excel reader
        /// </summary>
        /// <returns></returns>
        private IExcelDataReader GetExcelReader()
        {
            switch (_parameters.Type)
            {
                case ExcelDocumentType.Xml:
                    return ExcelReaderFactory.CreateOpenXmlReader(File.OpenRead(_parameters.Database));
                case ExcelDocumentType.Binary:
                    return ExcelReaderFactory.CreateBinaryReader(File.OpenRead(_parameters.Database));
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        /// <summary>
        /// Imports excel data to storage and opens connection
        /// </summary>
        private void OpenInternal()
        {
            var database = _storage.GetDatabaseName(_parameters.Database, _storegeDir);

            // Drop existing database, if ForceStorageReload is set
            if (_parameters.ForceStorageReload)
                _storage.DropDatabase(database, _storegeDir);

            if (!_storage.DatabaseExists(database, _storegeDir) ||
                _parameters.ForceStorageReload)
            {
                try
                {
                    _storage.CreateDatabase(database, _storegeDir);

                    using (
                        var tempStorageConnection =
                            _storage.GetConnection(
                                _storage.GetConnectionString(database, "", _parameters.Password)))
                    {
                        tempStorageConnection.Open();

                        var sourceReader = GetExcelReader();

                        _storage.ImportData(sourceReader, _parameters, tempStorageConnection);

                        tempStorageConnection.Close();
                    }
                }
                catch (Exception ex)
                {
                    // delete unitialized storage database
                    _storage.DropDatabase(database, _storegeDir);

                    throw new ExcelException(ex, "Cannot initialize storage for '{0}'", _parameters.Database);
                }

            }

            _storageConnection = _storage.GetConnection(
                _storage.GetConnectionString(database, "", _parameters.Password));
            _storageConnection.Open();

            // change connection state
            _storageInitialized = true;
            _state = ConnectionState.Open;
        }

        #endregion

        #region IDbConnection

        public override void Open()
        {
            CheckDisposed();

            lock (_syncLock)
            {
                OpenInternal();
            }
        }

        public override void Close()
        {
            CheckDisposed();

            lock (_syncLock)
            {
                try
                {
                    _storageConnection?.Close();
                }
                catch (Exception)
                {
                    // ignore
                }
                finally
                {
                    _storageInitialized = false;
                    _storageConnection = null;
                    _state = ConnectionState.Closed;
                }
            }
        }

        public override void ChangeDatabase(string databaseName)
        {
            CheckDisposed();

            throw new NotImplementedException();
        }

        protected override DbCommand CreateDbCommand()
        {
            CheckDisposed();

            lock (_syncLock)
            {
                if (_storageInitialized)
                {
                    return (DbCommand)_storageConnection.CreateCommand();
                }
            }

            throw new ExcelException("Internal storage was not initialied");
        }

        protected override DbTransaction BeginDbTransaction(IsolationLevel isolationLevel)
        {
            CheckDisposed();

            lock (_syncLock)
            {
                if (_storageInitialized)
                {
                    return (DbTransaction)_storageConnection.BeginTransaction(isolationLevel);
                }
            }

            throw new ExcelException("Internal storage was not initialied");
        }

        #region Properties 

        public override string ConnectionString
        {
            get { return _connectionString; }
            set
            {
                CheckDisposed();

                if (value == null)
                    throw new ArgumentException();

                if (State != ConnectionState.Closed)
                    throw new InvalidOperationException();

                InitSettings(value);
            }
        }

        public override string Database
        {
            get
            {
                CheckDisposed();

                return _parameters.Database;
            }
        }

        public override string DataSource
        {
            get
            {
                CheckDisposed();

                return Path.GetFileNameWithoutExtension(_parameters.Database);
            }
        }

        /// <summary>
        /// Gets Excel connector assembly version
        /// </summary>
        public override string ServerVersion => Assembly.GetExecutingAssembly().GetName().Version.ToString();

        /// <summary>
        /// Gets connection state
        /// </summary>
        public override ConnectionState State => _state;

        #endregion

        #endregion

        #region IDisposable

        /// <summary>
        /// Checks instance status
        /// </summary>
        private void CheckDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException(typeof(ExcelConnection).Name);
        }

        /// <summary>
        /// Disposes and finalizes connection
        /// </summary>
        public new void Dispose()
        {
            base.Dispose();
        }

        protected override void Dispose(bool disposing)
        {
            _disposing = true;

            try
            {
                if (_disposed)
                    return;

                Close();
            }
            finally
            {
                base.Dispose(disposing);
                _disposed = true;
            }
        }

        #endregion
    }
}
