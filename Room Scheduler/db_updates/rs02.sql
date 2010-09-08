-- MySQL Administrator dump 1.4
--
-- ------------------------------------------------------
-- Server version	5.0.45-community-nt


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;


--
-- Create schema room_scheduler
--

CREATE DATABASE IF NOT EXISTS room_scheduler;
USE room_scheduler;

--
-- Definition of table `rooms`
--

DROP TABLE IF EXISTS `rooms`;
CREATE TABLE `rooms` (
  `id` int(11) NOT NULL auto_increment,
  `name` varchar(100) default NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=utf8;

--
-- Dumping data for table `rooms`
--

/*!40000 ALTER TABLE `rooms` DISABLE KEYS */;
INSERT INTO `rooms` (`id`,`name`) VALUES 
 (1,'Operating Room'),
 (2,'Emergency Room'),
 (4,'OnCall Room');
/*!40000 ALTER TABLE `rooms` ENABLE KEYS */;


--
-- Definition of table `schedules`
--

DROP TABLE IF EXISTS `schedules`;
CREATE TABLE `schedules` (
  `id` int(11) NOT NULL auto_increment,
  `day` date NOT NULL,
  `start_time` time NOT NULL,
  `end_time` time NOT NULL,
  `room_id` int(11) NOT NULL,
  `subject_id` int(11) NOT NULL,
  `section_id` int(11) NOT NULL,
  PRIMARY KEY  (`id`),
  KEY `room_id` (`room_id`),
  KEY `subject_id` (`subject_id`),
  KEY `section_id` (`section_id`),
  CONSTRAINT `schedules_ibfk_1` FOREIGN KEY (`room_id`) REFERENCES `rooms` (`id`) ON DELETE CASCADE,
  CONSTRAINT `schedules_ibfk_2` FOREIGN KEY (`subject_id`) REFERENCES `subjects` (`id`) ON DELETE CASCADE,
  CONSTRAINT `schedules_ibfk_3` FOREIGN KEY (`section_id`) REFERENCES `sections` (`id`) ON DELETE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=utf8;

--
-- Dumping data for table `schedules`
--

/*!40000 ALTER TABLE `schedules` DISABLE KEYS */;
INSERT INTO `schedules` (`id`,`day`,`start_time`,`end_time`,`room_id`,`subject_id`,`section_id`) VALUES 
 (1,'2010-08-28','02:39:00','18:24:00',1,1,3),
 (3,'2010-09-01','18:00:00','01:00:00',1,1,4);
/*!40000 ALTER TABLE `schedules` ENABLE KEYS */;


--
-- Definition of table `section_users`
--

DROP TABLE IF EXISTS `section_users`;
CREATE TABLE `section_users` (
  `id` int(11) NOT NULL auto_increment,
  `section_id` int(11) default NULL,
  `user_id` int(11) default NULL,
  PRIMARY KEY  (`id`),
  KEY `user_id` (`user_id`),
  KEY `section_id` (`section_id`),
  CONSTRAINT `section_users_ibfk_1` FOREIGN KEY (`user_id`) REFERENCES `users` (`id`) ON DELETE CASCADE,
  CONSTRAINT `section_users_ibfk_2` FOREIGN KEY (`section_id`) REFERENCES `sections` (`id`) ON DELETE CASCADE
) ENGINE=InnoDB AUTO_INCREMENT=12 DEFAULT CHARSET=utf8;

--
-- Dumping data for table `section_users`
--

/*!40000 ALTER TABLE `section_users` DISABLE KEYS */;
INSERT INTO `section_users` (`id`,`section_id`,`user_id`) VALUES 
 (3,2,3),
 (9,1,1),
 (10,1,2),
 (11,1,7);
/*!40000 ALTER TABLE `section_users` ENABLE KEYS */;


--
-- Definition of table `sections`
--

DROP TABLE IF EXISTS `sections`;
CREATE TABLE `sections` (
  `id` int(11) NOT NULL auto_increment,
  `name` varchar(100) default NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=11 DEFAULT CHARSET=utf8;

--
-- Dumping data for table `sections`
--

/*!40000 ALTER TABLE `sections` DISABLE KEYS */;
INSERT INTO `sections` (`id`,`name`) VALUES 
 (1,'General Surgery  - Bailey'),
 (2,'Pediatrics'),
 (3,'Plastics'),
 (4,'Cardio Nueristics'),
 (5,'Nuero '),
 (6,'Orthod'),
 (8,'diagnostics'),
 (9,'whatever'),
 (10,'Test');
/*!40000 ALTER TABLE `sections` ENABLE KEYS */;


--
-- Definition of table `subjects`
--

DROP TABLE IF EXISTS `subjects`;
CREATE TABLE `subjects` (
  `id` int(11) NOT NULL auto_increment,
  `name` varchar(100) default NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=7 DEFAULT CHARSET=utf8;

--
-- Dumping data for table `subjects`
--

/*!40000 ALTER TABLE `subjects` DISABLE KEYS */;
INSERT INTO `subjects` (`id`,`name`) VALUES 
 (1,'Anatomy eeh'),
 (2,'Programming'),
 (3,'Dance and Arts'),
 (6,'Maturity');
/*!40000 ALTER TABLE `subjects` ENABLE KEYS */;


--
-- Definition of table `users`
--

DROP TABLE IF EXISTS `users`;
CREATE TABLE `users` (
  `id` int(11) NOT NULL auto_increment,
  `last_name` varchar(100) default NULL,
  `first_name` varchar(100) default NULL,
  `middle_name` varchar(100) default NULL,
  `type` enum('Instructor','Student') default NULL,
  PRIMARY KEY  (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=utf8;

--
-- Dumping data for table `users`
--

/*!40000 ALTER TABLE `users` DISABLE KEYS */;
INSERT INTO `users` (`id`,`last_name`,`first_name`,`middle_name`,`type`) VALUES 
 (1,'Grey','Meredith','McDreamy','Student'),
 (2,'Yeng','Christina','Hunt','Student'),
 (3,'Stevens','Isabelle','','Student'),
 (4,'Omalley','Goerge','','Student'),
 (7,'Kariv','Alex','','Student');
/*!40000 ALTER TABLE `users` ENABLE KEYS */;




/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
