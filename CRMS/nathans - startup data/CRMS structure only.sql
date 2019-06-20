CREATE DATABASE  IF NOT EXISTS `crms_db` /*!40100 DEFAULT CHARACTER SET latin1 */;
USE `crms_db`;
-- MySQL dump 10.13  Distrib 5.7.12, for Win64 (x86_64)
--
-- Host: 127.0.0.1    Database: crms_db
-- ------------------------------------------------------
-- Server version	5.1.67-community

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `antennas`
--

DROP TABLE IF EXISTS `antennas`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `antennas` (
  `antenna` varchar(45) NOT NULL,
  `manufacturer` varchar(45) DEFAULT NULL,
  `900` int(1) DEFAULT '0',
  `1800` int(1) DEFAULT '0',
  `2100` int(1) DEFAULT '0',
  `edt_min` int(11) DEFAULT '0',
  `edt_max` int(11) DEFAULT '0',
  `has_mdt` int(11) DEFAULT '1',
  `hbw` double DEFAULT '0',
  `vbw` double DEFAULT '0',
  `gain_dbi` double DEFAULT '0',
  `dual_beam` int(1) DEFAULT '0',
  `comment` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`antenna`),
  UNIQUE KEY `antenna_UNIQUE` (`antenna`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `cr_common`
--

DROP TABLE IF EXISTS `cr_common`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `cr_common` (
  `cr_id` varchar(100) NOT NULL,
  `technology` varchar(10) DEFAULT NULL,
  `cr_objective` varchar(100) DEFAULT NULL,
  `team` varchar(50) DEFAULT NULL,
  `region` varchar(50) DEFAULT NULL,
  `cr_type` text,
  `node_types` text,
  `issue_description` text,
  `expected_impact` text,
  `risk` text,
  `requester` varchar(200) DEFAULT NULL,
  `approver` varchar(200) DEFAULT NULL,
  `execution_coordinator` varchar(200) DEFAULT NULL,
  `executors` text,
  `open_date` datetime DEFAULT NULL,
  `approval_date` datetime DEFAULT NULL,
  `planned_execution_date` datetime DEFAULT NULL,
  `execution_date` datetime DEFAULT NULL,
  `closed_date` datetime DEFAULT NULL,
  `cr_status` varchar(50) DEFAULT NULL,
  `last_activity` varchar(50) DEFAULT NULL,
  `last_activity_date` datetime DEFAULT NULL,
  `last_nag_date` datetime DEFAULT NULL,
  `cr_lifetime` int(11) DEFAULT NULL,
  `resubmission_lifetime` int(11) DEFAULT NULL,
  `resubmission_nag_period` int(11) DEFAULT NULL,
  `missing_req_attach_lifetime` int(11) DEFAULT NULL,
  `missing_req_attach_nag_period` int(11) DEFAULT NULL,
  `approval_lifetime` int(11) DEFAULT NULL,
  `approval_nag_period` int(11) DEFAULT NULL,
  `execution_planning_lifetime` int(11) DEFAULT NULL,
  `execution_planning_nag_period` int(11) DEFAULT NULL,
  `execution_lifetime` int(11) DEFAULT NULL,
  `execution_nag_period` int(11) DEFAULT NULL,
  `missing_ex_attach_lifetime` int(11) DEFAULT NULL,
  `missing_ex_attach_nag_period` int(11) DEFAULT NULL,
  `review_lifetime` int(11) DEFAULT NULL,
  `review_nag_period` int(11) DEFAULT NULL,
  `msg_signature` text NOT NULL,
  `cc_list` text NOT NULL,
  `cr_type_short` varchar(5) NOT NULL,
  `cr_form_type` varchar(5) NOT NULL,
  PRIMARY KEY (`cr_id`),
  UNIQUE KEY `CR_ID_UNIQUE` (`cr_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `cr_data_oth`
--

DROP TABLE IF EXISTS `cr_data_oth`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `cr_data_oth` (
  `cr_sub_id` varchar(100) NOT NULL,
  `cr_type` varchar(50) DEFAULT NULL,
  `node_type` varchar(50) DEFAULT NULL,
  `node` varchar(50) DEFAULT NULL,
  `requester_comments` text,
  `execution_coordinator` varchar(100) DEFAULT NULL,
  `planned_execution_date` datetime DEFAULT NULL,
  `executor` varchar(100) DEFAULT NULL,
  `execution_status` varchar(20) DEFAULT NULL,
  `execution_date` datetime DEFAULT NULL,
  `executor_comments` text,
  PRIMARY KEY (`cr_sub_id`),
  UNIQUE KEY `cr_sub_id_UNIQUE` (`cr_sub_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='table is for re-eng and hw crs';
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `cr_data_prm`
--

DROP TABLE IF EXISTS `cr_data_prm`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `cr_data_prm` (
  `cr_sub_id` varchar(100) NOT NULL,
  `cr_type` varchar(50) DEFAULT NULL,
  `node_type` varchar(50) DEFAULT NULL,
  `node` varchar(50) DEFAULT NULL,
  `nbr_node` varchar(50) DEFAULT NULL,
  `parameter` varchar(50) DEFAULT NULL,
  `proposed_setting` text,
  `rollback_setting` text,
  `requester_comments` text,
  `execution_coordinator` varchar(100) DEFAULT NULL,
  `planned_execution_date` datetime DEFAULT NULL,
  `executor` varchar(100) DEFAULT NULL,
  `execution_status` varchar(20) DEFAULT NULL,
  `execution_date` datetime DEFAULT NULL,
  `executor_comments` text,
  PRIMARY KEY (`cr_sub_id`),
  UNIQUE KEY `cr_sub_id_UNIQUE` (`cr_sub_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='This is the log table for all CR events';
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `cr_data_rfb`
--

DROP TABLE IF EXISTS `cr_data_rfb`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `cr_data_rfb` (
  `cr_sub_id` varchar(100) NOT NULL,
  `cr_type` varchar(50) DEFAULT NULL,
  `node_type` varchar(50) DEFAULT NULL,
  `node` varchar(50) DEFAULT NULL,
  `cur_az` varchar(6) DEFAULT NULL,
  `cur_mdt` varchar(6) DEFAULT NULL,
  `cur_edt` varchar(6) DEFAULT NULL,
  `pro_az` varchar(6) DEFAULT NULL,
  `pro_mdt` varchar(6) DEFAULT NULL,
  `pro_edt` varchar(6) DEFAULT NULL,
  `requester_comments` text,
  `execution_coordinator` varchar(100) DEFAULT NULL,
  `planned_execution_date` datetime DEFAULT NULL,
  `executor` varchar(100) DEFAULT NULL,
  `act_az` varchar(6) DEFAULT NULL,
  `act_mdt` varchar(6) DEFAULT NULL,
  `act_edt` varchar(6) DEFAULT NULL,
  `fin_az` varchar(6) DEFAULT NULL,
  `fin_mdt` varchar(6) DEFAULT NULL,
  `fin_edt` varchar(6) DEFAULT NULL,
  `fin_ht` varchar(6) DEFAULT NULL,
  `fin_antenna` varchar(100) DEFAULT NULL,
  `fin_coax_len` varchar(6) DEFAULT NULL,
  `execution_status` varchar(20) DEFAULT NULL,
  `execution_date` datetime DEFAULT NULL,
  `executor_comments` text,
  PRIMARY KEY (`cr_sub_id`),
  UNIQUE KEY `cr_sub_id_UNIQUE` (`cr_sub_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='This is the log table for all CR events';
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `cr_obj_2g`
--

DROP TABLE IF EXISTS `cr_obj_2g`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `cr_obj_2g` (
  `cr_objective` varchar(100) NOT NULL,
  PRIMARY KEY (`cr_objective`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `cr_obj_3g`
--

DROP TABLE IF EXISTS `cr_obj_3g`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `cr_obj_3g` (
  `cr_objective` varchar(100) NOT NULL,
  PRIMARY KEY (`cr_objective`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `cr_obj_4g`
--

DROP TABLE IF EXISTS `cr_obj_4g`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `cr_obj_4g` (
  `cr_objective` varchar(100) NOT NULL,
  PRIMARY KEY (`cr_objective`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `cr_types`
--

DROP TABLE IF EXISTS `cr_types`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `cr_types` (
  `cr_type` varchar(45) NOT NULL,
  `cr_type_short` varchar(45) NOT NULL,
  `cr_form_type` varchar(45) NOT NULL,
  `approval_required` varchar(1) NOT NULL DEFAULT '1',
  `verification_required` varchar(1) NOT NULL DEFAULT '1',
  `cr_lifetime` int(11) NOT NULL DEFAULT '2000',
  `resubmission_lifetime` int(11) DEFAULT '160',
  `resubmission_nag_period` int(11) NOT NULL DEFAULT '160',
  `missing_req_attach_lifetime` int(11) NOT NULL DEFAULT '24',
  `missing_req_attach_nag_period` int(11) NOT NULL DEFAULT '24',
  `approval_lifetime` int(11) NOT NULL DEFAULT '6',
  `approval_nag_period` int(11) NOT NULL DEFAULT '6',
  `execution_planning_lifetime` int(11) NOT NULL DEFAULT '6',
  `execution_planning_nag_period` int(11) NOT NULL DEFAULT '6',
  `execution_lifetime` int(11) NOT NULL DEFAULT '72',
  `execution_nag_period` int(11) NOT NULL DEFAULT '72',
  `missing_ex_attach_lifetime` int(11) NOT NULL DEFAULT '24',
  `missing_ex_attach_nag_period` int(11) NOT NULL DEFAULT '24',
  `review_lifetime` int(11) NOT NULL DEFAULT '6',
  `review_nag_period` int(11) NOT NULL DEFAULT '6',
  PRIMARY KEY (`cr_type`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `geo`
--

DROP TABLE IF EXISTS `geo`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `geo` (
  `regency` varchar(45) NOT NULL,
  `province` varchar(45) NOT NULL,
  `province_short` varchar(45) NOT NULL,
  PRIMARY KEY (`regency`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `log`
--

DROP TABLE IF EXISTS `log`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `log` (
  `date` datetime DEFAULT NULL,
  `cr_id` varchar(45) DEFAULT NULL,
  `event` text NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1 COMMENT='This is a log for all activities in CRMS';
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `node_types`
--

DROP TABLE IF EXISTS `node_types`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `node_types` (
  `node_type` varchar(45) NOT NULL,
  `technology` varchar(45) NOT NULL,
  PRIMARY KEY (`node_type`),
  UNIQUE KEY `node_type_UNIQUE` (`node_type`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `parameters`
--

DROP TABLE IF EXISTS `parameters`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `parameters` (
  `location` varchar(10) NOT NULL,
  `mo` varchar(200) NOT NULL,
  `parameter` varchar(200) NOT NULL,
  `change_level` int(11) DEFAULT '0',
  `pattern` varchar(45) DEFAULT '.*',
  PRIMARY KEY (`location`,`mo`,`parameter`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `people`
--

DROP TABLE IF EXISTS `people`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `people` (
  `name` varchar(200) NOT NULL,
  `email` varchar(200) NOT NULL,
  `requester` int(11) NOT NULL DEFAULT '0',
  `approver` int(11) NOT NULL DEFAULT '0',
  `rfb_ex_coord` int(11) NOT NULL DEFAULT '0',
  `rfr_ex_coord` int(11) NOT NULL DEFAULT '0',
  `hdw_ex_coord` int(11) NOT NULL DEFAULT '0',
  `prm_ex_coord` int(11) NOT NULL DEFAULT '0',
  `executor` int(11) NOT NULL DEFAULT '0',
  `query` int(11) NOT NULL DEFAULT '0',
  `anyquery` int(11) NOT NULL DEFAULT '0',
  `administrator` int(11) NOT NULL DEFAULT '0',
  PRIMARY KEY (`name`,`email`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `risk`
--

DROP TABLE IF EXISTS `risk`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `risk` (
  `risk` varchar(10) NOT NULL DEFAULT 'Low',
  PRIMARY KEY (`risk`),
  UNIQUE KEY `risk_UNIQUE` (`risk`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `state_control`
--

DROP TABLE IF EXISTS `state_control`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `state_control` (
  `id` int(11) NOT NULL,
  `state` varchar(50) NOT NULL,
  `allowed_transitions` varchar(1000) NOT NULL,
  `can_cancel` int(11) NOT NULL DEFAULT '1',
  PRIMARY KEY (`state`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `teams`
--

DROP TABLE IF EXISTS `teams`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `teams` (
  `team` varchar(45) NOT NULL,
  `team_short` varchar(45) NOT NULL,
  PRIMARY KEY (`team`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Table structure for table `tech`
--

DROP TABLE IF EXISTS `tech`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `tech` (
  `tech` varchar(45) NOT NULL,
  PRIMARY KEY (`tech`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
/*!40101 SET character_set_client = @saved_cs_client */;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2017-01-08 19:23:48
