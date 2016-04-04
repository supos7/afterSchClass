-- 방과후 만대초

CREATE TABLE student (
   id       integer        PRIMARY KEY AUTOINCREMENT,
   name     varchar(64)    NOT NULL,
   year     smallint       NOT NULL,
   grade    smallint       NOT NULL,
   class    smallint       NOT NULL,
   odr      smallint       NOT NULL,
   sex      varchar(4),
   code     varchar(4),
   UNIQUE(year,grade,class,odr)
);

CREATE TABLE afterSchoolClass (
   id       integer        PRIMARY KEY AUTOINCREMENT,
   cname    varchar(128)   NOT NULL UNIQUE,
   year     smallint       NOT NULL,
   month    smallint       NOT NULL,
   tuition  integer,
   mcost    integer,
   mfee     integer,
   UNIQUE(cname,year,month)
);

CREATE TABLE classStu (
   id       integer        PRIMARY KEY AUTOINCREMENT,
   classId  integer        REFERENCES afterSchoolClass(id) ON DELETE CASCADE ON UPDATE CASCADE NOT NULL,
   stuId    integer        REFERENCES student(id) ON DELETE CASCADE ON UPDATE CASCADE NOT NULL,
   tuition  integer,
   mcost    integer,
   mfee     integer,
   code     varchar(4),
   UNIQUE(classId,stuId)
);


CREATE INDEX idx_yearName_stu ON student(year,name);
CREATE INDEX idx_yearTerm_cls ON afterSchoolClass(year,month);
CREATE INDEX idx_stuId_clsStu ON classStu(stuId);


CREATE VIEW afterSchStu AS
   SELECT student.id AS stuId,
      student.name AS name,
      student.year AS year,
      student.grade AS grade,
      student.class AS class,
      student.odr AS odr,
      afterSchoolClass.id AS classId,
      afterSchoolClass.cname AS cname,
      afterSchoolClass.month AS month,
      classStu.tuition AS tuition,
      classStu.mcost AS mcost,
      classStu.mfee AS mfee,
      classStu.code AS code
   FROM student,afterSchoolClass,classStu
   WHERE classStu.classId = afterSchoolClass.id AND classStu.stuId = student.id;
