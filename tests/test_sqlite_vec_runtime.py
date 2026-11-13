"""Actual Windows sqlite-vec extension smoke test; vectors are synthetic."""
import sqlite3
from contextlib import closing

import pytest
import sqlite_vec


def load(db):
    db.enable_load_extension(True)
    try:
        sqlite_vec.load(db)
    finally:
        db.enable_load_extension(False)


def test_real_extension_knn_persistence_rollback_and_disabled_extension_loading(tmp_path):
    path = tmp_path / 'index.sqlite3'
    with closing(sqlite3.connect(path)) as db:
        load(db)
        assert db.execute('select vec_version()').fetchone()[0].startswith('v0.1.')
        assert db.execute('pragma journal_mode=delete').fetchone()[0] == 'delete'
        db.execute('create virtual table vectors using vec0(embedding float[2] distance_metric=cosine)')
        db.execute('insert into vectors(rowid,embedding) values(?,?)',(1,sqlite_vec.serialize_float32([1,0])))
        db.execute('insert into vectors(rowid,embedding) values(?,?)',(2,sqlite_vec.serialize_float32([0,1])))
        db.commit()
        with pytest.raises(sqlite3.OperationalError,match='not authorized'):
            db.execute("select load_extension('fixture-not-an-extension')")
        db.execute('delete from vectors where rowid=1')
        db.rollback()
        assert db.execute('select count(*) from vectors').fetchone()[0] == 2
    with closing(sqlite3.connect(path.as_uri()+'?mode=ro',uri=True)) as db:
        load(db)
        rows = db.execute('select rowid,distance from vectors where embedding match ? and k=? order by distance',
                          (sqlite_vec.serialize_float32([0.99,0.01]),2)).fetchall()
        assert [row[0] for row in rows] == [1,2]
        assert rows[0][1] < 0.001 and rows[1][1] > 0.98
    assert sorted(item.name for item in tmp_path.iterdir()) == ['index.sqlite3']
