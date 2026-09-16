-- ==============================================================================
-- SEBN-TN Enterprise Maintenance Suite — MySQL / MariaDB Production Schema
-- Engine: InnoDB | Charset: utf8mb4 | Collation: utf8mb4_unicode_ci
-- Total Tables: 26
-- ==============================================================================

SET FOREIGN_KEY_CHECKS = 0;

-- 1. USERS
CREATE TABLE IF NOT EXISTS `users` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `name` VARCHAR(255) NOT NULL,
    `matricule` VARCHAR(100) NULL,
    `username` VARCHAR(191) NOT NULL UNIQUE,
    `password_hash` VARCHAR(255) NOT NULL,
    `role` VARCHAR(50) NOT NULL DEFAULT 'TECHNICIAN',
    `shift` VARCHAR(20) DEFAULT 'A',
    `is_active` TINYINT(1) DEFAULT 1,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    `last_login` DATETIME NULL,
    `technician_level` VARCHAR(100) NULL,
    `photo` VARCHAR(255) NULL,
    INDEX `idx_users_username` (`username`),
    INDEX `idx_users_matricule` (`matricule`),
    INDEX `idx_users_role` (`role`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 2. MACHINES
CREATE TABLE IF NOT EXISTS `machines` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `group_name` VARCHAR(100) NOT NULL DEFAULT '',
    `machine_id` VARCHAR(100) NOT NULL UNIQUE,
    `machine_name` VARCHAR(255) NOT NULL DEFAULT '',
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    INDEX `idx_machines_id` (`machine_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 3. INTERVENTIONS
CREATE TABLE IF NOT EXISTS `interventions` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `code` VARCHAR(100) NOT NULL UNIQUE,
    `machine_id` VARCHAR(100) NOT NULL,
    `group_name` VARCHAR(100) NOT NULL DEFAULT '',
    `fault_description` TEXT NOT NULL,
    `code_asp` VARCHAR(100) NOT NULL DEFAULT '',
    `start_time` VARCHAR(50) NOT NULL,
    `end_time` VARCHAR(50) NULL,
    `downtime_minutes` DOUBLE NOT NULL DEFAULT 0,
    `technician_name` VARCHAR(255) NOT NULL DEFAULT '',
    `technician_mat` VARCHAR(100) NOT NULL DEFAULT '',
    `shift` VARCHAR(20) NOT NULL DEFAULT 'A',
    `status` VARCHAR(50) NOT NULL DEFAULT 'OPEN',
    `priority` VARCHAR(50) NOT NULL DEFAULT 'MEDIUM',
    `category` VARCHAR(50) NOT NULL DEFAULT 'GEN',
    `remarks` TEXT NOT NULL,
    `checklist_results` LONGTEXT NOT NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `closed_at` VARCHAR(50) NULL,
    INDEX `idx_int_machine` (`machine_id`),
    INDEX `idx_int_status` (`status`),
    INDEX `idx_int_group` (`group_name`),
    INDEX `idx_int_created` (`created_at`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 4. MACHINE PLANS
CREATE TABLE IF NOT EXISTS `machine_plans` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `machine_id` VARCHAR(100) NOT NULL,
    `group_name` VARCHAR(100) NOT NULL,
    `plan_description` TEXT NOT NULL,
    `target_date` VARCHAR(50) NULL,
    `status` VARCHAR(50) NOT NULL DEFAULT 'PENDING',
    `created_by` VARCHAR(255) NOT NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `sent_email` INT NOT NULL DEFAULT 0,
    `file_path` VARCHAR(500) NULL,
    INDEX `idx_mp_machine` (`machine_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 5. ASP CODES
CREATE TABLE IF NOT EXISTS `asp_codes` (
    `code` VARCHAR(100) PRIMARY KEY,
    `description` TEXT NOT NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 6. SETTINGS
CREATE TABLE IF NOT EXISTS `settings` (
    `key` VARCHAR(191) PRIMARY KEY,
    `value` LONGTEXT NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 7. REPORTS
CREATE TABLE IF NOT EXISTS `reports` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `filename` VARCHAR(255) NOT NULL,
    `filepath` VARCHAR(500) NOT NULL,
    `type` VARCHAR(50) NOT NULL DEFAULT 'MANUAL',
    `sent_email` INT NOT NULL DEFAULT 0,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 8. EXAMS
CREATE TABLE IF NOT EXISTS `exams` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `title` VARCHAR(255) NOT NULL,
    `description` TEXT NULL,
    `technician_level` VARCHAR(100) NOT NULL,
    `level_percentage` INT NOT NULL,
    `source_file` VARCHAR(500) NULL,
    `duration` INT NOT NULL DEFAULT 30,
    `number_of_questions` INT DEFAULT 0,
    `passing_score` INT DEFAULT 75,
    `status` VARCHAR(50) DEFAULT 'draft',
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 9. EXAM QUESTIONS
CREATE TABLE IF NOT EXISTS `exam_questions` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `exam_id` INT NOT NULL,
    `question_number` INT NOT NULL,
    `question_text` LONGTEXT NOT NULL,
    `question_type` VARCHAR(50) DEFAULT 'qcm',
    `image` VARCHAR(500) NULL,
    `available_labels` LONGTEXT NULL,
    `boxes_json` LONGTEXT NULL,
    INDEX `idx_eq_exam` (`exam_id`),
    CONSTRAINT `fk_eq_exam` FOREIGN KEY (`exam_id`) REFERENCES `exams`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 10. EXAM ANSWERS
CREATE TABLE IF NOT EXISTS `exam_answers` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `question_id` INT NOT NULL,
    `answer_text` LONGTEXT NOT NULL,
    `is_correct` INT DEFAULT 0,
    INDEX `idx_ea_question` (`question_id`),
    CONSTRAINT `fk_ea_question` FOREIGN KEY (`question_id`) REFERENCES `exam_questions`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 11. EXAM ATTEMPTS
CREATE TABLE IF NOT EXISTS `exam_attempts` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `exam_id` INT NOT NULL,
    `user_id` INT NOT NULL,
    `started_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `submitted_at` DATETIME NULL,
    `score` INT NULL,
    `percentage` DOUBLE NULL,
    `correct_answers` INT NULL,
    `incorrect_answers` INT NULL,
    `unanswered` INT NULL,
    `status` VARCHAR(50) DEFAULT 'running',
    INDEX `idx_ea_exam` (`exam_id`),
    INDEX `idx_ea_user` (`user_id`),
    CONSTRAINT `fk_eat_exam` FOREIGN KEY (`exam_id`) REFERENCES `exams`(`id`) ON DELETE CASCADE,
    CONSTRAINT `fk_eat_user` FOREIGN KEY (`user_id`) REFERENCES `users`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 12. EXAM ATTEMPT ANSWERS
CREATE TABLE IF NOT EXISTS `exam_attempt_answers` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `attempt_id` INT NOT NULL,
    `question_id` INT NOT NULL,
    `selected_answer_id` INT NULL,
    `answer_text` LONGTEXT NULL,
    INDEX `idx_eaa_attempt` (`attempt_id`),
    INDEX `idx_eaa_question` (`question_id`),
    CONSTRAINT `fk_eaa_attempt` FOREIGN KEY (`attempt_id`) REFERENCES `exam_attempts`(`id`) ON DELETE CASCADE,
    CONSTRAINT `fk_eaa_question` FOREIGN KEY (`question_id`) REFERENCES `exam_questions`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 13. EXAM QUESTION IMAGES
CREATE TABLE IF NOT EXISTS `exam_question_images` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `question_id` INT NOT NULL,
    `filename` VARCHAR(500) NOT NULL,
    `sort_order` INT DEFAULT 0,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    INDEX `idx_eqi_question` (`question_id`),
    CONSTRAINT `fk_eqi_question` FOREIGN KEY (`question_id`) REFERENCES `exam_questions`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 14. LEVEL CHANGE LOG
CREATE TABLE IF NOT EXISTS `level_change_log` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `user_id` INT NOT NULL,
    `previous_level` VARCHAR(100) NULL,
    `new_level` VARCHAR(100) NOT NULL,
    `changed_by` VARCHAR(255) NOT NULL,
    `change_type` VARCHAR(50) DEFAULT 'EXAM',
    `reason` TEXT NULL,
    `exam_attempt_id` INT NULL,
    `changed_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    INDEX `idx_lcl_user` (`user_id`),
    CONSTRAINT `fk_lcl_user` FOREIGN KEY (`user_id`) REFERENCES `users`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 15. PASSATIONS
CREATE TABLE IF NOT EXISTS `passations` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `user_id` INT NULL,
    `technician_matricule` VARCHAR(100) NULL,
    `technician_name` VARCHAR(255) NULL,
    `shift` VARCHAR(20) NOT NULL,
    `target_shift` VARCHAR(20) NULL,
    `zone_name` VARCHAR(100) NULL,
    `remarks` LONGTEXT NULL,
    `status` VARCHAR(50) DEFAULT 'Completed',
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 16. PASSATION QUESTIONS
CREATE TABLE IF NOT EXISTS `passation_questions` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `text` TEXT NOT NULL,
    `category` VARCHAR(100) NOT NULL DEFAULT 'Général',
    `type` VARCHAR(50) NOT NULL DEFAULT 'CHOICE',
    `options` LONGTEXT NULL,
    `is_active` INT DEFAULT 1,
    `sort_order` INT DEFAULT 0
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 17. PASSATION RESPONSES
CREATE TABLE IF NOT EXISTS `passation_responses` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `passation_id` INT NOT NULL,
    `question_id` INT NOT NULL,
    `answer` LONGTEXT NULL,
    INDEX `idx_pr_passation` (`passation_id`),
    CONSTRAINT `fk_pr_passation` FOREIGN KEY (`passation_id`) REFERENCES `passations`(`id`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 18. PASSATION SETTINGS
CREATE TABLE IF NOT EXISTS `passation_settings` (
    `key` VARCHAR(191) PRIMARY KEY,
    `value` LONGTEXT NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 19. EBM SETTINGS
CREATE TABLE IF NOT EXISTS `ebm_settings` (
    `key` VARCHAR(191) PRIMARY KEY,
    `value` LONGTEXT NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 20. EBM ACTION PLANS
CREATE TABLE IF NOT EXISTS `ebm_action_plans` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `title` VARCHAR(255) NOT NULL,
    `description` LONGTEXT NULL,
    `responsible` VARCHAR(255) NULL,
    `due_date` VARCHAR(50) NULL,
    `priority` VARCHAR(50) DEFAULT 'Medium',
    `status` VARCHAR(50) DEFAULT 'Pending',
    `file_path` VARCHAR(500) NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 21. EBM VALIDATIONS
CREATE TABLE IF NOT EXISTS `ebm_validations` (
    `item_id` VARCHAR(191) PRIMARY KEY,
    `status` VARCHAR(50) NOT NULL,
    `validated_by` VARCHAR(255) NULL,
    `comment` LONGTEXT NULL,
    `validated_at` DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 22. DOCUMENTS
CREATE TABLE IF NOT EXISTS `documents` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `display_name` VARCHAR(255) NOT NULL,
    `file_name` VARCHAR(255) NOT NULL DEFAULT '',
    `file_type` VARCHAR(100) NOT NULL DEFAULT '',
    `storage_path` VARCHAR(500) NOT NULL DEFAULT '',
    `is_active` INT NOT NULL DEFAULT 1,
    `display_order` INT NOT NULL DEFAULT 0,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX `idx_docs_active` (`is_active`),
    INDEX `idx_docs_order` (`display_order`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 23. CHECKLIST DEFINITIONS
CREATE TABLE IF NOT EXISTS `checklist_definitions` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `checklist_code` VARCHAR(100) UNIQUE NOT NULL,
    `name` VARCHAR(255) NOT NULL,
    `target_type` VARCHAR(50) NOT NULL DEFAULT 'PPE',
    `equipment_pattern` VARCHAR(100) NOT NULL DEFAULT 'ALL',
    `description` TEXT NULL,
    `created_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `updated_at` DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 24. CHECKLIST VERSIONS
CREATE TABLE IF NOT EXISTS `checklist_versions` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `definition_id` INT NOT NULL,
    `version_number` INT NOT NULL,
    `status` VARCHAR(50) NOT NULL DEFAULT 'ACTIVE',
    `original_filename` VARCHAR(255) NOT NULL,
    `stored_filename` VARCHAR(255) NOT NULL,
    `file_sha256` VARCHAR(100) NULL,
    `item_count` INT NOT NULL DEFAULT 0,
    `imported_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `imported_by` VARCHAR(255) DEFAULT 'admin',
    `change_summary` TEXT NULL,
    UNIQUE KEY `uq_def_version` (`definition_id`, `version_number`),
    INDEX `idx_cv_definition` (`definition_id`),
    CONSTRAINT `fk_cv_definition` FOREIGN KEY (`definition_id`) REFERENCES `checklist_definitions`(`id`) ON DELETE RESTRICT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 25. CHECKLIST ITEMS
CREATE TABLE IF NOT EXISTS `checklist_items` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `version_id` INT NOT NULL,
    `item_number` INT NOT NULL,
    `section_header` VARCHAR(255) NULL,
    `description` TEXT NOT NULL,
    `method` TEXT NULL,
    `control_type` VARCHAR(50) DEFAULT 'OK_NOK_NA',
    `icon` VARCHAR(100) DEFAULT 'fa-clipboard-check',
    `row_index` INT NULL,
    INDEX `idx_ci_version` (`version_id`),
    CONSTRAINT `fk_ci_version` FOREIGN KEY (`version_id`) REFERENCES `checklist_versions`(`id`) ON DELETE RESTRICT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- 26. CHECKLIST EXECUTIONS
CREATE TABLE IF NOT EXISTS `checklist_executions` (
    `id` INT AUTO_INCREMENT PRIMARY KEY,
    `version_id` INT NULL,
    `equipment` VARCHAR(100) NOT NULL,
    `task_type` VARCHAR(100) NOT NULL,
    `sheet` VARCHAR(100) NULL,
    `week` VARCHAR(50) NULL,
    `month` VARCHAR(50) NULL,
    `technician_name` VARCHAR(255) NOT NULL,
    `technician_matricule` VARCHAR(100) NULL,
    `shift` VARCHAR(20) NULL,
    `executed_at` DATETIME DEFAULT CURRENT_TIMESTAMP,
    `status` VARCHAR(50) DEFAULT 'COMPLETED',
    `answers_json` LONGTEXT NOT NULL,
    `filled_excel_path` VARCHAR(500) NULL,
    INDEX `idx_ce_version` (`version_id`),
    CONSTRAINT `fk_ce_version` FOREIGN KEY (`version_id`) REFERENCES `checklist_versions`(`id`) ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

SET FOREIGN_KEY_CHECKS = 1;
