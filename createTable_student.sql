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
   cname    varchar(128)   NOT NULL,
   year     smallint       NOT NULL,
   month    smallint       NOT NULL,
   tuition  integer,
   mcost    integer,
   mfee     integer,
   UNIQUE(cname,year,month)
);

CREATE TABLE classStu (
   id       integer        PRIMARY KEY AUTOINCREMENT,
   classId  integer        REFERENCES afterSchoolClass(id) ON DELETE CASCADE NOT NULL,
   stuId    integer        REFERENCES student(id) ON DELETE CASCADE NOT NULL,
   month    integer        NOT NULL,
   tuition  integer        DEFAULT 0,
   mcost    integer        DEFAULT 0,
   mfee     integer        DEFAULT 0,
   code     varchar(4),
   tuit_pay char(2),
   mcos_pay char(2),
   mfee_pay char(2),
   quitNew  char(2),
   UNIQUE(classId,stuId,month)
);


CREATE INDEX idx_yearName_stu ON student(year,name);
CREATE INDEX idx_yearTerm_cls ON afterSchoolClass(year,month);
CREATE INDEX idx_stuId_clsStu ON classStu(classId,stuId,month);


CREATE VIEW afterSchStu AS
   SELECT student.id AS stuId,
      student.name AS name,
      student.year AS year,
      student.grade AS grade,
      student.class AS class,
      student.odr AS odr,
      student.code AS scode,
      afterSchoolClass.id AS classId,
      afterSchoolClass.cname AS cname,
      afterSchoolClass.year AS cyear,
      afterSchoolClass.month AS cmonth,
      classStu.id AS classStuId,
      classStu.month AS month,
      classStu.tuition AS tuition,
      classStu.mcost AS mcost,
      classStu.mfee AS mfee,
      classStu.code AS code,
      classStu.tuit_pay AS tuit_pay,
      classStu.mcos_pay AS mcos_pay,
      classStu.mfee_pay AS mfee_pay,
      classStu.quitNew AS quitNew
   FROM student,afterSchoolClass,classStu
   WHERE classStu.classId = afterSchoolClass.id AND classStu.stuId = student.id;


CREATE TRIGGER delete_class_stu_id
BEFORE DELETE ON student
FOR EACH ROW BEGIN
   DELETE FROM classStu WHERE stuId = OLD.id;
END;
