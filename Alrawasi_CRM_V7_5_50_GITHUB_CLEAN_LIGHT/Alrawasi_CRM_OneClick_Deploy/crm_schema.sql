-- Structural schema only: NO customer data.
CREATE TABLE companies(id INTEGER PRIMARY KEY AUTOINCREMENT,source_id TEXT UNIQUE,name_ar TEXT,name_en TEXT,sector TEXT,subcategory TEXT,district TEXT,city TEXT,address TEXT,phone_landline TEXT,phone_mobile TEXT,whatsapp TEXT,email TEXT,website TEXT,cr_number TEXT,social_media TEXT,source TEXT,source_url TEXT,source_verified INTEGER,marketing_status TEXT,marketing_notes TEXT,recommended_services TEXT,created_at TEXT,lat REAL,lng REAL,assigned_user_id INTEGER,priority TEXT DEFAULT 'normal',updated_at TEXT DEFAULT CURRENT_TIMESTAMP, classification_status TEXT, classification_grade TEXT, muqawil_member_no TEXT, business_area TEXT DEFAULT 'contracting', data_source TEXT, imported_at TEXT, latitude REAL, longitude REAL, logo_url TEXT, google_place_id TEXT, phone_source TEXT, phone_verified_at TEXT, phone_match_name TEXT, phone_match_address TEXT, muqawil_profile_url TEXT, muqawil_profile_status TEXT, crm_number INTEGER, muqawil_key TEXT, muqawil_internal_id TEXT, muqawil_contractor_type TEXT, company_size TEXT, account_status TEXT, training_hours INTEGER, main_contractor_count INTEGER, subcontractor_count INTEGER, source_region_id TEXT, "membership_type" TEXT, "membership_since" TEXT, "muqawil_email" TEXT, "muqawil_phone_landline" TEXT, "muqawil_phone_mobile" TEXT, "muqawil_address" TEXT, "muqawil_city" TEXT, "muqawil_region" TEXT, "muqawil_enriched_at" TEXT);
CREATE TABLE interactions(id INTEGER PRIMARY KEY AUTOINCREMENT,company_id INTEGER,user_id INTEGER,channel TEXT,outcome TEXT,notes TEXT,contacted_at TEXT DEFAULT CURRENT_TIMESTAMP,next_followup_at TEXT, user_name_snapshot TEXT, remote_event_id TEXT);
CREATE TABLE templates(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT,channel TEXT,subject TEXT,body TEXT);
CREATE TABLE users(id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT,email TEXT UNIQUE,password_hash TEXT,role TEXT,active INTEGER DEFAULT 1, specialty text default 'general', whatsapp_template text, appearance_default TEXT);
CREATE INDEX idx_companies_assigned ON companies(assigned_user_id);
CREATE UNIQUE INDEX idx_companies_crm_number ON companies(crm_number);
CREATE UNIQUE INDEX idx_companies_muqawil_key ON companies(muqawil_key) WHERE muqawil_key IS NOT NULL;
CREATE INDEX idx_companies_muqawil_member ON companies(muqawil_member_no) WHERE muqawil_member_no IS NOT NULL;
CREATE INDEX idx_company_city ON companies(city);
CREATE INDEX idx_company_name ON companies(name_ar,name_en);
CREATE INDEX idx_company_status ON companies(marketing_status);
CREATE INDEX idx_interactions_company ON interactions(company_id);
CREATE UNIQUE INDEX idx_interactions_remote_event
ON interactions(remote_event_id)
WHERE remote_event_id IS NOT NULL AND remote_event_id<>'';
