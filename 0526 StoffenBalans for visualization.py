# ==============================================================================
# REMAS — Complete Integrated Calculation Pipeline 
# ==============================================================================
# Modules:
#   1.1  FarmManureCalculator              Manure Volume & Net/Gross N Excretion
#   1.2  NitrogenPartitioningCalculator    MUN-based UUN / FN split
#   2.1  VEMRequirementCalculator          Energy requirements (kVEM/yr)
#   2.2  VEMAllocationCalculator           VEM allocation to feed sources
#   2.3  NitrogenIntakeCalculator          DS, N, CP, VRE intake
#   2.4  NitrogenExcretionCalculatorVCRE   VCRE N retention & excretion
#   2.5  NitrogenPartitioningVCRE          VCRE-based UUN / FN split
#   3.1    MineralizationCalculator          Net Mineralization (MUN & VCRE)
#   3.2  CorrectedTANCalculator            Corrected TAN (MUN & VCRE)
#   4.1    EmissionCalculator                Stable / Storage / Grazing emissions
#   4.2  LandApplicationCalculator         N limits, Manure & Fertiliser
#.  4.3  ApplicationEmissionCalculator      N residue, application emisisons
# ==============================================================================
#
# INPUT:  InputREMAS.xlsx  (sheet 'Main input')
# OUTPUT: Output_REMAS_Complete.xlsx
#
# ==============================================================================

import numpy as np
import pandas as pd
import sys
import traceback

pd.options.mode.chained_assignment = None  # silence SettingWithCopyWarning

def ensure_numeric(df, col, default=0.0):
    if col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.replace(',', '.', regex=False).str.strip()
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(default)
    else:
        df[col] = default
    return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 1.1 — Manure Volume & N Excretion
# ══════════════════════════════════════════════════════════════════════════════
class FarmManureCalculator:
    CF_slurry = 0.86 #Net to gross excretion
    CF_solid = 0.61 #Net to gross excretion
    
    def __init__(self, input_filepath, sheet_name='Main input'):
        self.input_filepath = input_filepath
        self.sheet_name = sheet_name
        self.df = None

    def __init__(self, filepath, sheet_name='Main input'):
        self.filepath = filepath
        self.sheet_name = sheet_name

    def load_and_clean(self):
        print(f"--- Loading data from: {self.filepath} ---")
        try:
            df = pd.read_excel(self.filepath, sheet_name=self.sheet_name)
        except Exception as e:
            raise ValueError(f"Failed to load Excel file: {e}")

        df.columns = df.columns.str.strip()

        # Added all new inventory and nutrient override columns to numeric_targets
        numeric_targets = [
            'Nr_koe', 'Nr_pink', 'Nr_kalf', 'MilkYield', 'Fat%', 'Pro%', 'MilkUerum',
            'slurry_koe', 'solid_koe', 'slurry_kalf', 'solid_kalf', 'slurry_pink', 'solid_pink',
            'volume_slurry_koe', 'volume_solid_koe', 'volume_slurry_kalf', 'volume_solid_kalf',
            'volume_slurry_pink', 'volume_solid_pink', 'Kg_WholeMilk_Kalf',
            'Kg_KunstMelk_Kalf_0', 'Kg_KunstMelk_Kalf_B', 'Kg_KunstMelk_Kalf_t',
            'slurry%_koe', 'slurry%_kalf', 'slurry%_pink', 'NatureGL%', 'Ha_Grass', 'Ha_Mais', 
            'GD_Limited_Koe', 'GD_Combi_Koe', 'GD_Unlimited_Koe', 'GD_Unlimited_Kalf', 'GD_Unlimited_Pink',
            'GH_Koe', 'GH_Kalf', 'GH_Pink', 'DS_Freshcut',
            'DS_GrassSilage_0', 'DS_GrassSilage_B', 'DS_GrassSilage_S', 'DS_GrassSilage_t',
            'DS_MaizeSilage_0', 'DS_MaizeSilage_B', 'DS_MaizeSilage_S', 'DS_MaizeSilage_t',
            'DS_OtherSilage1_0', 'DS_OtherSilage1_B', 'DS_OtherSilage1_S', 'DS_OtherSilage1_t',
            'DS_OtherSilage2_0', 'DS_OtherSilage2_B', 'DS_OtherSilage2_S', 'DS_OtherSilage2_t',
            'DS_OtherSilage3_0', 'DS_OtherSilage3_B', 'DS_OtherSilage3_S', 'DS_OtherSilage3_t',
            'DS_Byproducts1_0', 'DS_Byproducts1_B', 'DS_Byproducts1_S', 'DS_Byproducts1_t',
            'DS_Byproducts2_0', 'DS_Byproducts2_B', 'DS_Byproducts2_S', 'DS_Byproducts2_t',
            'DS_Byproducts3_0', 'DS_Byproducts3_B', 'DS_Byproducts3_S', 'DS_Byproducts3_t',
            'Kg_conc1_0', 'Kg_conc1_B', 'Kg_conc1_t', 'Kg_conc2_0', 'Kg_conc2_B', 'Kg_conc2_t',
            'Kg_conc3_0', 'Kg_conc3_B', 'Kg_conc3_t',
            'N_fgrass', 'N_GrassSilage', 'N_MaizeSilage', 'N_OtherSilage1', 'N_OtherSilage2', 'N_OtherSilage3',
            'N_Byproducts1', 'N_Byproducts2', 'N_Byproducts3', 'N_Concentrate1', 'N_Concentrate2', 'N_Concentrate3',
            'VEM_fgrass', 'VEM_GrassSilage', 'VEM_MaizeSilage', 'VEM_OtherSilage1', 'VEM_OtherSilage2', 'VEM_OtherSilage3',
            'VEM_Byproducts1', 'VEM_Byproducts2', 'VEM_Byproducts3', 'VEM_Concentrate1', 'VEM_Concentrate2', 'VEM_Concentrate3',
            'DS_GrassSilage_Harvest_Input', 'DS_MaizeSilage_Harvest_Input'
        ]

        print("--- Cleaning Data (Fixing Comma/Dot issues) ---")
        for col in numeric_targets:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace(',', '.', regex=False)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            else:
                if 'slurry%' in col: df[col] = 1.0
                else: df[col] = 0.0

        return df

    def calculate_totals(self):
        if self.df is None: raise ValueError("Call load_and_clean() first")
        df = self.df
        sk, fk = df['slurry%_koe'], 1.0 - df['slurry%_koe']
        sc, fc = df['slurry%_kalf'], 1.0 - df['slurry%_kalf']
        sp, fp = df['slurry%_pink'], 1.0 - df['slurry%_pink']
        df['vol_koe_slurry'] = df['Nr_koe']*sk*df['volume_slurry_koe']
        df['vol_koe_solid'] = df['Nr_koe']*fk*df['volume_solid_koe']
        df['vol_kalf_slurry'] = df['Nr_kalf']*sc*df['volume_slurry_kalf']
        df['vol_kalf_solid'] = df['Nr_kalf']*fc*df['volume_solid_kalf']
        df['vol_pink_slurry'] = df['Nr_pink']*sp*df['volume_slurry_pink']
        df['vol_pink_solid'] = df['Nr_pink']*fp*df['volume_solid_pink']
        df['Total_Manure_Volume_m3'] = (df['vol_koe_slurry']+df['vol_koe_solid']+
            df['vol_kalf_slurry']+df['vol_kalf_solid']+df['vol_pink_slurry']+df['vol_pink_solid'])
        df['net_n_koe_slurry'] = df['Nr_koe']*sk*df['slurry_koe']
        df['net_n_koe_solid'] = df['Nr_koe']*fk*df['solid_koe']
        df['net_n_cows'] = df['net_n_koe_slurry']+df['net_n_koe_solid']
        df['gross_n_cows_mun'] = df['net_n_koe_slurry']/self.CF_slurry+df['net_n_koe_solid']/self.CF_solid
        
        df['net_n_kalf_slurry'] = df['Nr_kalf']*sc*df['slurry_kalf']
        df['net_n_kalf_solid'] = df['Nr_kalf']*fc*df['solid_kalf']
        df['net_n_calves'] = df['net_n_kalf_slurry']+df['net_n_kalf_solid']
        df['gross_n_calves_mun'] = df['net_n_kalf_slurry']/self.CF_slurry+df['net_n_kalf_solid']/self.CF_solid
        
        df['net_n_pink_slurry'] = df['Nr_pink']*sp*df['slurry_pink']
        df['net_n_pink_solid'] = df['Nr_pink']*fp*df['solid_pink']
        df['net_n_heifers'] = df['net_n_pink_slurry']+df['net_n_pink_solid']
        df['gross_n_heifers_mun'] = df['net_n_pink_slurry']/self.CF_slurry+df['net_n_pink_solid']/self.CF_solid
        
        df['Total_Net_Nitrogen_kg'] = df['net_n_cows']+df['net_n_calves']+df['net_n_heifers']
        df['Total_Nitrogen_Excretion_MUN'] = df['gross_n_cows_mun']+df['gross_n_calves_mun']+df['gross_n_heifers_mun']
      
        print(f"  ✓ 1.1 done — Farm1 Gross N: {df['Total_Nitrogen_Excretion_MUN'].iloc[0]:.1f} kg")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 1.2 — N Partitioning (MUN)
# ══════════════════════════════════════════════════════════════════════════════
class NitrogenPartitioningCalculator:
    
    def __init__(self, df): self.df = df
    def calculate_partitioning(self):
        print(f"\n{'='*70}\nMODULE 1.2: N Partitioning (MUN)\n{'='*70}")
        df = self.df
        df = ensure_numeric(df, 'MilkUerum', 0.0)
        df['MUN_value'] = df['MilkUerum']*(28.0/60.0)
        df['uun_per_cow_g_day'] = 16.7+13.0+12.03*df['MUN_value']
        df['uun_total_cows_kg'] = df['uun_per_cow_g_day']*365.0*df['Nr_koe']/1000.0
        df['fn_total_cows_kg'] = df['gross_n_cows_mun']-df['uun_total_cows_kg']
        df['cow_uun_ratio'] = np.where(df['gross_n_cows_mun']>0, df['uun_total_cows_kg']/df['gross_n_cows_mun'], 0.0)
        df['uun_total_kalf_kg'] = df['gross_n_calves_mun']*df['cow_uun_ratio']
        df['fn_total_kalf_kg'] = df['gross_n_calves_mun']-df['uun_total_kalf_kg']
        df['uun_total_pink_kg'] = df['gross_n_heifers_mun']*df['cow_uun_ratio']
        df['fn_total_pink_kg'] = df['gross_n_heifers_mun']-df['uun_total_pink_kg']
        df['Farm_Total_UUN_kg'] = df['uun_total_cows_kg']+df['uun_total_kalf_kg']+df['uun_total_pink_kg']
        df['Farm_Total_FN_kg'] = df['fn_total_cows_kg']+df['fn_total_kalf_kg']+df['fn_total_pink_kg']
        print(f"  ✓ 1.2 done — Farm1 UUN: {df['Farm_Total_UUN_kg'].iloc[0]:.1f}")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.1 — VEM Requirements
# ══════════════════════════════════════════════════════════════════════════════

class VEMRequirementCalculator:
    
    # 1. ADDED: __init__ method to accept the dataframe
    def __init__(self, df):
        self.df = df.copy()
        
    # 2. ADDED: _breed method to prevent the 'self._breed()' call from failing
    def _breed(self):
        if 'breed_factor' not in self.df.columns:
            self.df['breed_factor'] = 1.0  # Default to 1.0 (Holstein) if not provided
        if 'avg_weight' not in self.df.columns:
            self.df['avg_weight'] = 650.0  # Default cow weight if not provided

    def calculate_requirements(self):
        print(f"\n{'='*70}\nMODULE 2.1: VEM Requirements\n{'='*70}")
        
        columns_to_ensure = ['MilkYield','Fat%','Pro%','Nr_koe','Nr_pink','Nr_kalf',
                             'GD_Limited_Koe','GD_Combi_Koe','GD_Unlimited_Koe',
                             'GD_Unlimited_Kalf','GD_Unlimited_Pink']
        for col in columns_to_ensure:
            self.df = ensure_numeric(self.df, col, 0.0)
            
        self._breed()
        
        # Define constants for lactation and dry periods (consistent with legacy logic)
        lactation_days = 326.0
        dry_days = 39.0 
        
        # Extract frequently used variables
        fat_pct = self.df['Fat%']
        protein_pct = self.df['Pro%']
        breed_factor = self.df['breed_factor']
        
        # ---------------------------------------------------------
        # Kalf (Calves < 1 year)
        # ---------------------------------------------------------
        vem_kalf_growth = 1323.0 * breed_factor
        vem_kalf_exercise = self.df['GD_Unlimited_Kalf'] * 0.346 * breed_factor
        
        # 1.02 accounts for feed processing/losses (voederverliezen/opslag)
        self.df['vem_req_kalf_per_head_yr'] = (vem_kalf_growth + vem_kalf_exercise) * 1.02
        self.df['Total_VEM_Kalf_Farm'] = self.df['vem_req_kalf_per_head_yr'] * self.df['Nr_kalf']
        
        # ---------------------------------------------------------
        # Pink (Heifers > 1 year)
        # ---------------------------------------------------------
        vem_pink_growth = 2259.0 * breed_factor
        vem_pink_exercise = self.df['GD_Unlimited_Pink'] * 0.784 * breed_factor
        vem_pink_processing = 115.9 * breed_factor # Processing/Pregnancy maintenance
        
        self.df['vem_req_pink_per_head_yr'] = (vem_pink_growth + vem_pink_exercise + vem_pink_processing) * 1.02
        self.df['Total_VEM_Pink_Farm'] = self.df['vem_req_pink_per_head_yr'] * self.df['Nr_pink']
        
        # ---------------------------------------------------------
        # Cow (Dairy Cows)
        # ---------------------------------------------------------
        # 1. Milk Production Requirements
        # FPCM: Fat and Protein Corrected Milk
        fpcm_yearly = (0.337 + 0.116 * fat_pct + 0.06 * protein_pct) * self.df['MilkYield'] * 365.0 
        fpcm_daily_lactating = fpcm_yearly / lactation_days
        
        # Conversion factor for lactating cows based on daily yield
        conversion_factor_lactation = 1.0 + (fpcm_daily_lactating - 15.0) * 0.00165
        
        self.df['vem_cow_milk_yr'] = (442.0 * fpcm_daily_lactating * conversion_factor_lactation / 1000.0) * lactation_days
        
        # 2. Maintenance Requirements (Metabolic Weight)
        metabolic_weight = np.power(self.df['avg_weight'], 0.75)
        
        # Maintenance during lactation
        vem_maint_lactation = 42.4 * metabolic_weight * conversion_factor_lactation * lactation_days / 1000.0
        
        # Maintenance during dry period (FPCM = 0)
        conversion_factor_dry = 1.0 + (-15.0 * 0.00165)
        vem_maint_dry = 42.4 * metabolic_weight * conversion_factor_dry * dry_days / 1000.0
        
        self.df['vem_cow_maint_yr'] = vem_maint_lactation + vem_maint_dry
        
        # 3. Exercise Requirements (Grazing)
        # Calculate combined grazing days factor based on grazing system
        grazing_days_combined = (self.df['GD_Limited_Koe'] * 0.419 + 
                                 self.df['GD_Combi_Koe'] * 0.419 + 
                                 self.df['GD_Unlimited_Koe'] * 0.560)
                                 
        self.df['vem_cow_exercise_yr'] = 201.0 + grazing_days_combined * (lactation_days / 365.0) * breed_factor
        
        # 4. Growth/Youth and Pregnancy Requirements
        self.df['vem_cow_youth_yr'] = 102.0 * breed_factor
        self.df['vem_cow_preg_yr'] = 194.0 * breed_factor
        
        vem_cow_subtotal = (self.df['vem_cow_milk_yr']+self.df['vem_cow_maint_yr']+
              self.df['vem_cow_exercise_yr']+self.df['vem_cow_youth_yr']+self.df['vem_cow_preg_yr'])
        
        self.df['vem_req_cow_per_head_yr'] = vem_cow_subtotal*1.02
        
        self.df['Total_VEM_Cow_Farm'] = self.df['vem_req_cow_per_head_yr']*self.df['Nr_koe']
        self.df['Total_VEM_Requirement_Farm_kVEM'] = (
            self.df['Total_VEM_Cow_Farm']+self.df['Total_VEM_Kalf_Farm']+self.df['Total_VEM_Pink_Farm'])
        print(f"  ✓ 2.1 done — Farm1 VEM: {self.df['Total_VEM_Requirement_Farm_kVEM'].iloc[0]:.0f}")
        return self.df
    
# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.2 — VEM Allocation (Restored BEX Ratio Logic + Dynamic Inventory)
# ══════════════════════════════════════════════════════════════════════════════
class VEMAllocationCalculator:
    VEM_KuntsMelk=1500.0; VEM_Concentrate=940.0; Concentrate_DS_Conversion=0.876
    Loss_Milk=0.02; Loss_Kuntsmelk=0.02; Loss_Concentrate=0.02; Loss_Silage=0.05; Loss_Others=0.03 
    Pct_Grass_Kalf=0.75; Pct_Maize_Kalf=0.25; Pct_Grass_Pink=0.90; Pct_Maize_Pink=0.10
    
    def __init__(self, df): 
        self.df = df.copy()
        # 预设 PGK, PMK, PGP, PMP 用于 young stock 粗饲料分配比例
        self.PGK = self.Pct_Grass_Kalf
        self.PMK = self.Pct_Maize_Kalf
        self.PGP = self.Pct_Grass_Pink
        self.PMP = self.Pct_Maize_Pink
    
    @staticmethod
    def _soil_p(st):
        s = str(st).lower().strip()
        if 'klei' in s:   
            return {'yield_gs':10005, 'yield_fg':8893, 'yield_ms':17685, 'yield_nat_gs':6000, 'yield_nat_fg':5333, 'vem_gs_cult':960.0, 'vem_gs_nat':842.0, 'vem_fg_cult':940.0, 'vem_fg_nat':860.0, 'vem_maize':950.0}
        elif 'zand' in s: 
            return {'yield_gs':9360, 'yield_fg':8320, 'yield_ms':17773, 'yield_nat_gs':6000, 'yield_nat_fg':5333, 'vem_gs_cult':960.0, 'vem_gs_nat':842.0, 'vem_fg_cult':940.0, 'vem_fg_nat':860.0, 'vem_maize':950.0}
        elif 'veen' in s: 
            return {'yield_gs':9751, 'yield_fg':8668, 'yield_ms':16620, 'yield_nat_gs':6000, 'yield_nat_fg':5333, 'vem_gs_cult':957.0, 'vem_gs_nat':842.0, 'vem_fg_cult':937.0, 'vem_fg_nat':860.0, 'vem_maize':950.0} 
        else:
            return {'yield_gs':9700, 'yield_fg':9700, 'yield_ms':17300, 'yield_nat_gs':6000, 'yield_nat_fg':6000, 'vem_gs_cult':960.0, 'vem_gs_nat':842.0, 'vem_fg_cult':940.0, 'vem_fg_nat':860.0, 'vem_maize':950.0}
                
    def _milk_vem(self):
        fat_pct, protein_pct = self.df['Fat%'], self.df['Pro%']
        GE = 744.38+365.7*fat_pct+241.4*protein_pct; ME = 584.17+376.6*fat_pct*0.94+171.5*protein_pct*0.87
        Q = np.where(GE>0, ME/GE*100.0, 0.0)
        return (0.6*(1.0+0.004*(Q-57.0))*0.9752*ME)/6.9
        
    def run_allocation(self):
        print(f"\n{'='*70}\nMODULE 2.2: VEM Allocation (Dynamic Input Method)\n{'='*70}")
        
        # --- 1. Calculate Aggregated Inventories & Nets ---
        # Kunstmelk
        self.df['Kg_KunstMelk_Kalf_Net'] = self.df['Kg_KunstMelk_Kalf_0'] + self.df['Kg_KunstMelk_Kalf_B'] - self.df['Kg_KunstMelk_Kalf_t']
        self.df['Kg_KunstMelk_Kalf'] = np.where(self.df['Kg_WholeMilk_Kalf'] > 0, self.df['Kg_WholeMilk_Kalf'], self.df['Kg_KunstMelk_Kalf_Net']).clip(min=0)
        
        # Silage Inventory Dynamic: Opening + Bought - Sold - Closing
        self.df['DS_GS_InvNet'] = self.df['DS_GrassSilage_0'] + self.df['DS_GrassSilage_B'] - self.df['DS_GrassSilage_S'] - self.df['DS_GrassSilage_t']
        self.df['DS_MS_InvNet'] = self.df['DS_MaizeSilage_0'] + self.df['DS_MaizeSilage_B'] - self.df['DS_MaizeSilage_S'] - self.df['DS_MaizeSilage_t']
        
        # Sum up the 3 columns for Conc, Byprod, OtherSilage
        self.df['DS_Byproducts_Total'] = sum((self.df[f'DS_Byproducts{i}_0'] + self.df[f'DS_Byproducts{i}_B'] - self.df[f'DS_Byproducts{i}_S'] - self.df[f'DS_Byproducts{i}_t']) for i in [1,2,3]).clip(lower=0)
        self.df['DS_OtherSilage_Total'] = sum((self.df[f'DS_OtherSilage{i}_0'] + self.df[f'DS_OtherSilage{i}_B'] - self.df[f'DS_OtherSilage{i}_S'] - self.df[f'DS_OtherSilage{i}_t']) for i in [1,2,3]).clip(lower=0)
        
        self.df['DS_CutGrass_Total'] = self.df.get('DS_Freshcut', pd.Series(0, index=self.df.index))
        
        # --- 2. Soil logic and weighted VEM ---
        sp = self.df['Soil_Type'].apply(self._soil_p).tolist()
        for k in ('yield_gs','yield_fg','yield_ms','yield_nat_gs','yield_nat_fg',
                  'vem_gs_cult','vem_gs_nat','vem_fg_cult','vem_fg_nat','vem_maize'):
            self.df[k] = [d[k] for d in sp]
            
        pn = self.df['NatureGL%'].clip(0,100)
        self.df['vem_fg_weighted'] = ((100-pn)*self.df['vem_fg_cult'] + pn*self.df['vem_fg_nat'])/100
        self.df['vem_gs_weighted'] = ((100-pn)*self.df['vem_gs_cult'] + pn*self.df['vem_gs_nat'])/100
        self.df['yield_fg_weighted'] = ((100-pn)*self.df['yield_fg'] + pn*self.df['yield_nat_fg'])/100
        self.df['yield_gs_weighted'] = ((100-pn)*self.df['yield_gs'] + pn*self.df['yield_nat_gs'])/100
        
        # Overrides if user provided measured VEM
        self.df['vem_fg_weighted'] = np.where(self.df['VEM_fgrass'] > 0, self.df['VEM_fgrass'], self.df['vem_fg_weighted'])
        self.df['vem_gs_weighted'] = np.where(self.df['VEM_GrassSilage'] > 0, self.df['VEM_GrassSilage'], self.df['vem_gs_weighted'])
        self.df['vem_maize'] = np.where(self.df['VEM_MaizeSilage'] > 0, self.df['VEM_MaizeSilage'], self.df['vem_maize'])
        
        # --- 3. Calculate Total VEM Intake for the fixed/known feeds ---
        self.df['kVEM_Intake_Byproducts_Total'] = 0.0
        self.df['kVEM_Intake_OtherSilage_Total'] = 0.0
        self.df['kVEM_Intake_Conc_Total'] = 0.0
        
        for i in [1,2,3]:
            vb = self.df[f'VEM_Byproducts{i}'].replace(0, 950.0)
            vo = self.df[f'VEM_OtherSilage{i}'].replace(0, 950.0)
            vc = self.df[f'VEM_Concentrate{i}'].replace(0, self.VEM_Concentrate)
            
            ds_bp = (self.df[f'DS_Byproducts{i}_0'] + self.df[f'DS_Byproducts{i}_B'] - self.df[f'DS_Byproducts{i}_S'] - self.df[f'DS_Byproducts{i}_t']).clip(lower=0)
            ds_os = (self.df[f'DS_OtherSilage{i}_0'] + self.df[f'DS_OtherSilage{i}_B'] - self.df[f'DS_OtherSilage{i}_S'] - self.df[f'DS_OtherSilage{i}_t']).clip(lower=0)
            kg_cc = (self.df[f'Kg_conc{i}_0'] + self.df[f'Kg_conc{i}_B'] - self.df[f'Kg_conc{i}_t']).clip(lower=0)
            
            self.df['kVEM_Intake_Byproducts_Total'] += ds_bp * (1-self.Loss_Others) * vb / 1000
            self.df['kVEM_Intake_OtherSilage_Total'] += ds_os * (1-self.Loss_Others) * vo / 1000
            self.df['kVEM_Intake_Conc_Total'] += kg_cc * self.Concentrate_DS_Conversion * (1-self.Loss_Concentrate) * vc / 1000
            
        # --- 4. Young Stock Allocation ---
        vmd = self._milk_vem()
        self.df['kVEM_Intake_Milk_Kalf'] = self.df['Kg_WholeMilk_Kalf']*(1-self.Loss_Milk)*vmd/1000
        self.df['kVEM_Intake_KunstMelk_Kalf'] = self.df['Kg_KunstMelk_Kalf']*(1-self.Loss_Kuntsmelk)*self.VEM_KuntsMelk/1000
        
        rgk = (self.df['GD_Unlimited_Kalf']/365).clip(0,1)
        rgp = (self.df['GD_Unlimited_Pink']/365).clip(0,1)
        self.df['kVEM_Intake_Conc_Kalf'] = self.df['Total_VEM_Kalf_Farm']*(0.10*rgk+0.25*(1-rgk))
        self.df['kVEM_Intake_Conc_Pink'] = self.df['Total_VEM_Pink_Farm']*(0.00*rgp+0.05*(1-rgp))
        
        tk = rgk*(1323.0-101.2)+self.df['GD_Unlimited_Kalf']*0.346
        self.df['kVEM_Intake_FreshGrass_Kalf'] = self.df['Nr_kalf']*tk*0.9*self.df['breed_factor']*1.02
        tp = rgp*(2259.0+102.9)+self.df['GD_Unlimited_Pink']*0.784
        self.df['kVEM_Intake_FreshGrass_Pink'] = self.df['Nr_pink']*tp*self.df['breed_factor']*1.02
        
        rk = (self.df['Total_VEM_Kalf_Farm']-self.df['kVEM_Intake_Milk_Kalf']-self.df['kVEM_Intake_KunstMelk_Kalf']-self.df['kVEM_Intake_Conc_Kalf']-self.df['kVEM_Intake_FreshGrass_Kalf']).clip(lower=0)
        self.df['kVEM_Intake_GrassSilage_Kalf'] = rk*self.PGK
        self.df['kVEM_Intake_MaizeSilage_Kalf'] = rk*self.PMK
        
        rp = (self.df['Total_VEM_Pink_Farm']-self.df['kVEM_Intake_Conc_Pink']-self.df['kVEM_Intake_FreshGrass_Pink']).clip(lower=0)
        self.df['kVEM_Intake_GrassSilage_Pink'] = rp*self.PGP
        self.df['kVEM_Intake_MaizeSilage_Pink'] = rp*self.PMP
        
        self.df['kVEM_Intake_Conc_Cow'] = (self.df['kVEM_Intake_Conc_Total'] - self.df['kVEM_Intake_Conc_Kalf'] - self.df['kVEM_Intake_Conc_Pink']).clip(lower=0)

        # --- 5. The BEX Ratio Squeeze Pool for Roughage ---
        hrs = self.df['GH_Koe']; ir = np.where(hrs>2, 2.0+0.75*(hrs-2), hrs)
        tgd = self.df['GD_Limited_Koe']+self.df['GD_Combi_Koe']+self.df['GD_Unlimited_Koe']
        fy = (0.337+0.116*self.df['Fat%']+0.06*self.df['Pro%'])*self.df['MilkYield']*365
        mf = 1.0+((fy-9500*self.df['breed_factor'])/500)*0.02; dc = (365-39)/365
        
        # Total Fresh Grass Pool = Cow physical grazing + Young stock grazing + Cut Grass
        cgv = tgd*ir*(self.df['vem_fg_weighted']/1000)*dc*mf*self.df['breed_factor']*self.df['Nr_koe']
        kVEM_CutGrass = self.df['DS_CutGrass_Total'] * (1-self.Loss_Others) * (self.df['vem_fg_weighted']/1000)
        tgv = cgv + self.df['kVEM_Intake_FreshGrass_Pink'] + self.df['kVEM_Intake_FreshGrass_Kalf'] + kVEM_CutGrass
        
        # Land used for Fresh Grass
        tfm = tgv / (self.df['vem_fg_weighted']/1000).replace(0, np.nan)
        hfn = (tfm / self.df['yield_fg_weighted'].replace(0, np.nan)).fillna(0)
        hag = (self.df['Ha_Grass'] - hfn).clip(lower=0)
        
        # ─── 新增动态收成逻辑 (Dynamic Silage Harvest) ───
        # 理论产出 (Forfaitair)
        harvest_gs_forfaitair = hag * self.df['yield_gs_weighted']
        harvest_ms_forfaitair = self.df['Ha_Mais'] * self.df['yield_ms']
        
        # 判断用户选择的是 "eigen kuilwaarden" 还是 "Forfaitair"
        input_source = self.df.get('Input_source_for_silage', pd.Series('Forfaitair', index=self.df.index)).astype(str).str.lower()
        is_eigen = input_source.str.contains('eigen')
        
        # 根据选择决定最终的 Aanleeg (Harvest)
        harvest_gs_actual = np.where(is_eigen, self.df.get('DS_GrassSilage_Harvest_Input', 0.0), harvest_gs_forfaitair)
        harvest_ms_actual = np.where(is_eigen, self.df.get('DS_MaizeSilage_Harvest_Input', 0.0), harvest_ms_forfaitair)
        
        # Silage Pool = Harvest + (Opening - Closing) + (Bought - Sold)
        ds_gs_avail = (harvest_gs_actual + self.df['DS_GS_InvNet']).clip(lower=0)
        ds_ms_avail = (harvest_ms_actual + self.df['DS_MS_InvNet']).clip(lower=0)
        
        vgs = ds_gs_avail * (self.df['vem_gs_weighted']/1000)
        vms = ds_ms_avail * (self.df['vem_maize']/1000)
        
        # Calculate Ratios
        thg = tgv + vgs + vms
        mk = thg > 0
        rfg = np.where(mk, tgv/thg, 0)
        rgs = np.where(mk, vgs/thg, 0)
        rms = np.where(mk, vms/thg, 0)
        
        # --- 6. Calculate the remaining VEM gap for the farm ---
        hgr = (self.df['Total_VEM_Cow_Farm'] + self.df['Total_VEM_Kalf_Farm'] + self.df['Total_VEM_Pink_Farm']
               - self.df['kVEM_Intake_Conc_Total']
               - self.df['kVEM_Intake_Milk_Kalf'] - self.df['kVEM_Intake_KunstMelk_Kalf']
               - self.df['kVEM_Intake_Byproducts_Total'] - self.df['kVEM_Intake_OtherSilage_Total']).clip(lower=0)
        
        # --- 7. Apply the Ratio Squeeze and assign residual to Cows ---
        self.df['kVEM_Intake_FreshGrass_Cow'] = (hgr*rfg - self.df['kVEM_Intake_FreshGrass_Kalf'] - self.df['kVEM_Intake_FreshGrass_Pink']).clip(lower=0)
        self.df['kVEM_Intake_GrassSilage_Cow'] = (hgr*rgs - self.df['kVEM_Intake_GrassSilage_Kalf'] - self.df['kVEM_Intake_GrassSilage_Pink']).clip(lower=0)
        self.df['kVEM_Intake_MaizeSilage_Cow'] = (hgr*rms - self.df['kVEM_Intake_MaizeSilage_Kalf'] - self.df['kVEM_Intake_MaizeSilage_Pink']).clip(lower=0)
        
        # Record Byproducts & Other Silages explicitly to cows
        self.df['kVEM_Intake_Byproducts_Cow'] = self.df['kVEM_Intake_Byproducts_Total']
        self.df['kVEM_Intake_OtherSilage_Cow'] = self.df['kVEM_Intake_OtherSilage_Total']
        
        print(f"  ✓ 2.2 done")
        return self.df
    
# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.3 — N Intake (Restored to simple DS conversion)
# ══════════════════════════════════════════════════════════════════════════════
class NitrogenIntakeCalculator:
    FP=6.25; FD=6.38; RAS=40.0; NK=32.52; DSW=0.228; DSK=0.964
    def __init__(self, df): self.df = df.copy()
    
    def _soil_N(self, st):
        s = str(st).lower().strip()
        if 'klei' in s:   return {'ng':28.2,'nf':31.6,'nm':12.0,'nn_gs':24.7,'nn_fg':30.2} 
        elif 'zand' in s: return {'ng':28.6,'nf':32.0,'nm':12.0,'nn_gs':25.1,'nn_fg':30.2}
        elif 'veen' in s: return {'ng':27.8,'nf':31.1,'nm':12.0,'nn_gs':24.5,'nn_fg':30.2}
        else:             return {'ng':28.2,'nf':31.6,'nm':12.0,'nn_gs':24.7,'nn_fg':30.2}
        
    def _vre(self, ds, cpg, ft):
        cp = cpg; cps = np.where(cp>0, cp, 1.0)
        if ft=='grass_fresh': v = (0.931*cp-43.2)/cps
        elif ft=='grass_silage': v = (0.963*cp-38.3)/cps
        elif ft=='maize_silage': v = (0.969*cp+0.04*self.RAS-40.0)/cps
        elif ft=='concentrate': v = (88.7*(1.0-np.exp(-0.012*cp)))/100.0
        elif ft=='dairy': v = pd.Series(0.89, index=ds.index) if hasattr(ds, 'index') else 0.89
        else: v = 0.0
        v = np.clip(v, 0, 1)
        return ds*(v*cp)/1000.0
        
    def run_nutrient_calculation(self):
        print(f"\n{'='*70}\nMODULE 2.3: N Intake\n{'='*70}")
        sn = self.df['Soil_Type'].apply(self._soil_N).tolist()
        
        self.df['N_cont_fresh_soil'] = [d['nf'] for d in sn]
        self.df['N_cont_gs_soil'] = [d['ng'] for d in sn]
        self.df['N_cont_ms_soil'] = [d['nm'] for d in sn]
        self.df['N_cont_nat_gs'] = [d['nn_gs'] for d in sn]
        self.df['N_cont_nat_fg'] = [d['nn_fg'] for d in sn]
        
        pn = self.df['NatureGL%'].clip(0,100)
        self.df['N_cont_fresh_weighted'] = ((100-pn)*self.df['N_cont_fresh_soil'] + pn*self.df['N_cont_nat_fg'])/100
        self.df['N_cont_gs_weighted'] = ((100-pn)*self.df['N_cont_gs_soil'] + pn*self.df['N_cont_nat_gs'])/100
        
        # Apply N Overrides
        self.df['N_cont_fresh_weighted'] = np.where(self.df['N_fgrass'] > 0, self.df['N_fgrass'], self.df['N_cont_fresh_weighted'])
        self.df['N_cont_gs_weighted'] = np.where(self.df['N_GrassSilage'] > 0, self.df['N_GrassSilage'], self.df['N_cont_gs_weighted'])
        self.df['N_cont_ms_soil'] = np.where(self.df['N_MaizeSilage'] > 0, self.df['N_MaizeSilage'], self.df['N_cont_ms_soil'])

        vfgw = self.df['vem_fg_weighted'] 
        vgsw = self.df['vem_gs_weighted']
        vmc = self.df['vem_maize']
        
        n_milk_ds = (self.df['Pro%']*10.0/self.DSW)/self.FD
        
        # 1. Convert Allocated VEM back to Dry Matter (DS)
        for sfx,fg,gs,ms,cc in [
            ('cow','kVEM_Intake_FreshGrass_Cow','kVEM_Intake_GrassSilage_Cow','kVEM_Intake_MaizeSilage_Cow','kVEM_Intake_Conc_Cow'),
            ('pink','kVEM_Intake_FreshGrass_Pink','kVEM_Intake_GrassSilage_Pink','kVEM_Intake_MaizeSilage_Pink','kVEM_Intake_Conc_Pink'),
            ('kalf','kVEM_Intake_FreshGrass_Kalf','kVEM_Intake_GrassSilage_Kalf','kVEM_Intake_MaizeSilage_Kalf','kVEM_Intake_Conc_Kalf')]:
            
            self.df[f'DS_fresh_{sfx}'] = (self.df[fg]*1000/vfgw).fillna(0)
            self.df[f'DS_gs_{sfx}'] = (self.df[gs]*1000/vgsw).fillna(0)
            self.df[f'DS_ms_{sfx}'] = (self.df[ms]*1000/vmc.replace(0,np.nan)).fillna(0)
            
            # Weighted average of the 3 concentrates for accurate reverse DS conversion
            vc_avg = self.df[['VEM_Concentrate1', 'VEM_Concentrate2', 'VEM_Concentrate3']].replace(0, np.nan).mean(axis=1).fillna(940.0)
            self.df[f'DS_conc_{sfx}'] = (self.df[cc]*1000/vc_avg).fillna(0)
            
        self.df['DS_milk_kalf'] = self.df['Kg_WholeMilk_Kalf']*self.DSW
        self.df['DS_kunst_kalf'] = self.df['Kg_KunstMelk_Kalf']*self.DSK
        
        # 2. Add Byproducts & Other Silage back to Cow's DS
        self.df['DS_Total_Byproducts_Cow'] = self.df['DS_Byproducts_Total']
        self.df['DS_Total_OtherSilage_Cow'] = self.df['DS_OtherSilage_Total']
        
        # Averages for N contents
        self.df['N_Byproducts_Avg'] = self.df[['N_Byproducts1', 'N_Byproducts2', 'N_Byproducts3']].replace(0, 25.0).mean(axis=1)
        self.df['N_OtherSilage_Avg'] = self.df[['N_OtherSilage1', 'N_OtherSilage2', 'N_OtherSilage3']].replace(0, 20.0).mean(axis=1)
        self.df['N_Concentrate_Avg'] = self.df[['N_Concentrate1', 'N_Concentrate2', 'N_Concentrate3']].replace(0, 27.3).mean(axis=1)
        
        # 3. Calculate Final Nutrients (N, CP, VRE)
        for sfx in ('cow','pink','kalf'):
            lbl = sfx.capitalize()
            feeds = [(f'DS_fresh_{sfx}','N_cont_fresh_weighted','grass_fresh',False),
                     (f'DS_gs_{sfx}','N_cont_gs_weighted','grass_silage',False), 
                     (f'DS_ms_{sfx}','N_cont_ms_soil','maize_silage',False),
                     (f'DS_conc_{sfx}','N_Concentrate_Avg','concentrate',False)]
                     
            if sfx=='kalf':
                feeds.append(('DS_milk_kalf', n_milk_ds, 'dairy', True))
                feeds.append(('DS_kunst_kalf', self.NK, 'dairy', True))
                
            if sfx=='cow':
                # Add byproducts and other silage to Cow's nutrient load
                feeds.append(('DS_Total_Byproducts_Cow', 'N_Byproducts_Avg', 'concentrate', False)) 
                feeds.append(('DS_Total_OtherSilage_Cow', 'N_OtherSilage_Avg', 'grass_silage', False))
                
            tn=pd.Series(0.0,index=self.df.index); tcp=tn.copy(); tvre=tn.copy()
            for dc,nr,ft,dairy in feeds:
                ds = self.df[dc]; nc = self.df[nr] if isinstance(nr,str) else nr
                nf = self.FD if dairy else self.FP
                nk = ds*nc/1000; cpk = nk*nf; cpc = nc*nf
                vk = self._vre(ds, cpc, ft)
                tn+=nk; tcp+=cpk; tvre+=vk
                
            self.df[f'Total_N_Intake_{lbl}'] = tn
            self.df[f'Total_CP_Intake_{lbl}'] = tcp
            self.df[f'Total_VRE_Intake_{lbl}'] = tvre
            
        for l in ('Cow','Pink','Kalf'):
            self.df[f'VCRE_Factor_{l}'] = (self.df[f'Total_VRE_Intake_{l}']/self.df[f'Total_CP_Intake_{l}'].replace(0,np.nan)).fillna(0)
            
        tvf = sum(self.df[f'Total_VRE_Intake_{l}'] for l in ('Cow','Pink','Kalf'))
        tcf = sum(self.df[f'Total_CP_Intake_{l}'] for l in ('Cow','Pink','Kalf'))
        self.df['VCRE_Factor_Farm'] = (tvf/tcf.replace(0,np.nan)).fillna(0)
        self.df['Total_N_Intake_Farm'] = sum(self.df[f'Total_N_Intake_{l}'] for l in ('Cow','Pink','Kalf'))
        
        print(f"  ✓ 2.3 done — Farm1 N intake: {self.df['Total_N_Intake_Farm'].iloc[0]:.1f}")
        return self.df
    
# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.4 — N Excretion (VCRE)
# ══════════════════════════════════════════════════════════════════════════════
class NitrogenExcretionCalculatorVCRE:
    NK=29.4; NP=24.1; NV=23.1; NC=22.5
    RB=44/650; R1=320/650; RC=540/650
    BR=0.70; RR=0.27; FD=6.38
    def __init__(self, df): self.df = df.copy()
    def calculate_excretion(self):
        print(f"\n{'='*70}\nMODULE 2.4: N Excretion (VCRE)\n{'='*70}")
        df = self.df; wc = df['avg_weight']
        wb=wc*self.RB; wp=wc*self.R1; wv=wc*self.RC
        nb=wb*self.NK/1000; npk=wp*self.NP/1000; nv=wv*self.NV/1000; nc=wc*self.NC/1000
        ym = np.where(df['MilkYield']<100, 365, 1)
        tm = df['MilkYield']*ym*df['Nr_koe']
        rm = tm*df['Pro%']*10/self.FD/1000
        rf = nb*self.BR*df['Nr_koe']
        nih = self.RR*nv*df['Nr_koe']; noc = self.RR*nc*df['Nr_koe']
        df['N_Retention_Cow_VCRE'] = rm+rf+(noc-nih)
        gt = npk-nb; g1 = 0.36*df['breed_factor']
        t1 = gt*(0.376/0.407); t2 = (g1/2*24)*(0.031/0.407)
        ncr = np.divide(t1+t2, gt, out=np.ones_like(gt.values,dtype=float), where=gt.values!=0)
        df['N_Retention_Kalf_VCRE'] = gt*df['Nr_kalf']*ncr
        gha = (nv-npk)*(12/14)
        df['N_Retention_Pink_VCRE'] = (nb+gha)*df['Nr_pink']
        df['Total_N_Retention_Farm_VCRE'] = df['N_Retention_Cow_VCRE']+df['N_Retention_Kalf_VCRE']+df['N_Retention_Pink_VCRE']
        df['Total_N_Excretion_Cow_VCRE'] = df['Total_N_Intake_Cow']-df['N_Retention_Cow_VCRE']
        df['Total_N_Excretion_Kalf_VCRE'] = df['Total_N_Intake_Kalf']-df['N_Retention_Kalf_VCRE']
        df['Total_N_Excretion_Pink_VCRE'] = df['Total_N_Intake_Pink']-df['N_Retention_Pink_VCRE']
        df['Total_N_Excretion_Farm_VCRE'] = df['Total_N_Excretion_Cow_VCRE']+df['Total_N_Excretion_Kalf_VCRE']+df['Total_N_Excretion_Pink_VCRE']
        print(f"  ✓ 2.4 done — Farm1 excr: {df['Total_N_Excretion_Farm_VCRE'].iloc[0]:.1f}")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2.5 — N Partitioning (VCRE)
# ══════════════════════════════════════════════════════════════════════════════
class NitrogenPartitioningVCRE:
    DCF = 0.91  #converting to dairy cow
    def __init__(self, df): self.df = df.copy()
    def calculate_partitioning(self):
        print(f"\n{'='*70}\nMODULE 2.5: N Partitioning (VCRE)\n{'='*70}")
        df = self.df
        for co,lb in [('cow','Cow'),('pink','Pink'),('kalf','Kalf')]:
            dn = df[f'Total_N_Intake_{lb}']*df[f'VCRE_Factor_{lb}']*self.DCF
            uun = dn - df[f'N_Retention_{lb}_VCRE']
            fn = df[f'Total_N_Excretion_{lb}_VCRE'] - uun
            df[f'UUN_{lb}_VCRE'] = uun; df[f'FN_{lb}_VCRE'] = fn
        df['Total_UUN_Farm_VCRE'] = df['UUN_Cow_VCRE']+df['UUN_Kalf_VCRE']+df['UUN_Pink_VCRE']
        df['Total_FN_Farm_VCRE'] = df['FN_Cow_VCRE']+df['FN_Kalf_VCRE']+df['FN_Pink_VCRE']
        print(f"  ✓ 2.5 done — Farm1 VCRE UUN: {df['Total_UUN_Farm_VCRE'].iloc[0]:.1f}")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 3.1 — Mineralization
# ══════════════════════════════════════════════════════════════════════════════
class MineralizationCalculator:
    RS=0.10; RV=-0.25
    def __init__(self, df): self.df = df.copy()
    def calculate_mineralization(self):
        print(f"\n{'='*70}\nMODULE 3.1: Mineralization\n{'='*70}")
        df = self.df
        for an,fnc,uunc,sl,sfx in [
            ('Cow','fn_total_cows_kg','uun_total_cows_kg','slurry%_koe','Cow'),
            ('Kalf','fn_total_kalf_kg','uun_total_kalf_kg','slurry%_kalf','Kalf'),
            ('Pink','fn_total_pink_kg','uun_total_pink_kg','slurry%_pink','Pink')]:
            ns = df[fnc]*df[sl]*self.RS; nv = df[uunc]*(1-df[sl])*self.RV
            df[f'Net_Min_{sfx}_MUN'] = ns+nv
        df['Total_Net_Mineralization_MUN'] = df['Net_Min_Cow_MUN']+df['Net_Min_Kalf_MUN']+df['Net_Min_Pink_MUN']
        for an,fnc,uunc,sl,sfx in [
            ('Cow','FN_Cow_VCRE','UUN_Cow_VCRE','slurry%_koe','Cow'),
            ('Kalf','FN_Kalf_VCRE','UUN_Kalf_VCRE','slurry%_kalf','Kalf'),
            ('Pink','FN_Pink_VCRE','UUN_Pink_VCRE','slurry%_pink','Pink')]:
            ns = df[fnc]*df[sl]*self.RS; nv = df[uunc]*(1-df[sl])*self.RV
            df[f'Net_Min_{sfx}_VCRE'] = ns+nv
        df['Total_Net_Mineralization_VCRE'] = df['Net_Min_Cow_VCRE']+df['Net_Min_Kalf_VCRE']+df['Net_Min_Pink_VCRE']
        df['Mineralization_Diff'] = df['Total_Net_Mineralization_VCRE']-df['Total_Net_Mineralization_MUN']
        print(f"  ✓ 3.1 done — MUN: {df['Total_Net_Mineralization_MUN'].iloc[0]:.2f}, VCRE: {df['Total_Net_Mineralization_VCRE'].iloc[0]:.2f}")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 3.2 — Corrected TAN
# ══════════════════════════════════════════════════════════════════════════════
class CorrectedTANCalculator:
    def __init__(self, df): self.df = df.copy()
    def calculate_corrected_tan(self):
        print(f"\n{'='*70}\nMODULE 3.2: Corrected TAN\n{'='*70}")
        df = self.df
        for m,us,ms in [
            ('MUN',['uun_total_cows_kg','uun_total_kalf_kg','uun_total_pink_kg'],
                   ['Net_Min_Cow_MUN','Net_Min_Kalf_MUN','Net_Min_Pink_MUN']),
            ('VCRE',['UUN_Cow_VCRE','UUN_Kalf_VCRE','UUN_Pink_VCRE'],
                    ['Net_Min_Cow_VCRE','Net_Min_Kalf_VCRE','Net_Min_Pink_VCRE'])]:
            for an,uc,mc in zip(['Cow','Kalf','Pink'],us,ms):
                df[f'Corrected_TAN_{an}_{m}'] = df[uc]+df[mc]
            df[f'Total_Corrected_TAN_{m}'] = df[f'Corrected_TAN_Cow_{m}']+df[f'Corrected_TAN_Kalf_{m}']+df[f'Corrected_TAN_Pink_{m}']
        print(f"  ✓ 3.2 done — TAN MUN: {df['Total_Corrected_TAN_MUN'].iloc[0]:.1f}, VCRE: {df['Total_Corrected_TAN_VCRE'].iloc[0]:.1f}")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 4.1 — Emissions (Stable / Storage / Grazing)
# ══════════════════════════════════════════════════════════════════════════════
class EmissionCalculator:
    EFS=0.143; EFG=0.04; EFST=0.01; EFSO=0.024; EFVO=0.035; MSO=0.20; FN2N=14/17
    def __init__(self, df): self.df = df.copy()
    def calculate_emissions(self):
        print(f"\n{'='*70}\nMODULE 4.1: Emissions\n{'='*70}")
        df = self.df
        for c in ['GD_Limited_Koe','GD_Combi_Koe','GD_Unlimited_Koe','GH_Koe',
                   'GD_Unlimited_Pink','GH_Pink','GD_Unlimited_Kalf','GH_Kalf',
                   'slurry%_koe','slurry%_kalf','slurry%_pink']:
            df = ensure_numeric(df, c, 0.0)
        df['total_gd_cow'] = df['GD_Limited_Koe']+df['GD_Combi_Koe']+df['GD_Unlimited_Koe']
        cats = [
            {'n':'Cow','gd':'total_gd_cow','gh':'GH_Koe','sl':'slurry%_koe',
             'gm':'gross_n_cows_mun','um':'uun_total_cows_kg','tm':'Corrected_TAN_Cow_MUN',
             'gv':'Total_N_Excretion_Cow_VCRE','uv':'UUN_Cow_VCRE','tv':'Corrected_TAN_Cow_VCRE'},
            {'n':'Pink','gd':'GD_Unlimited_Pink','gh':'GH_Pink','sl':'slurry%_pink',
             'gm':'gross_n_heifers_mun','um':'uun_total_pink_kg','tm':'Corrected_TAN_Pink_MUN',
             'gv':'Total_N_Excretion_Pink_VCRE','uv':'UUN_Pink_VCRE','tv':'Corrected_TAN_Pink_VCRE'},
            {'n':'Kalf','gd':'GD_Unlimited_Kalf','gh':'GH_Kalf','sl':'slurry%_kalf',
             'gm':'gross_n_calves_mun','um':'uun_total_kalf_kg','tm':'Corrected_TAN_Kalf_MUN',
             'gv':'Total_N_Excretion_Kalf_VCRE','uv':'UUN_Kalf_VCRE','tv':'Corrected_TAN_Kalf_VCRE'}]
        for ct in cats:
            nm = ct['n']; gd = df[ct['gd']].fillna(0); gh = df[ct['gh']].fillna(0)
            sp = df[ct['sl']].fillna(1)
            gr = (0.0261*gh*(gd/365)).clip(0,1)
            epi = (1-(gd*gh)/(365*24)).clip(0,1)
            for method in ['mun','vcre']:
                M = 'MUN' if method=='mun' else 'VCRE'
                gk = 'gm' if method=='mun' else 'gv'
                uk = 'um' if method=='mun' else 'uv'
                tk = 'tm' if method=='mun' else 'tv'
                gn = df[ct[gk]].fillna(0); uun = df[ct[uk]].fillna(0); tan = df[ct[tk]].fillna(0)
                se = (tan*self.EFS*(1-gr))/self.FN2N
                df[f'Emission_Stable_{nm}_{M}'] = se
                ni = gn*epi
                no = ni*(sp*self.EFSO+(1-sp)*self.EFVO)
                nts = ((ni-no)/self.FN2N-se).clip(lower=0)
                #to add solid manure (AS)
                ste = nts*self.MSO*self.EFST
                df[f'Emission_Storage_{nm}_{M}'] = ste
                uo = uun*(1-epi)
                ge = uo*self.EFG/self.FN2N
                df[f'Emission_Grazing_{nm}_{M}'] = ge
        for M in ['MUN','VCRE']:
            for tp in ['Stable','Storage','Grazing']:
                df[f'Total_Emission_{tp}_{M}'] = sum(df[f'Emission_{tp}_{n}_{M}'] for n in ['Cow','Pink','Kalf'])
            df[f'Total_Emission_All_{M}'] = df[f'Total_Emission_Stable_{M}']+df[f'Total_Emission_Storage_{M}']+df[f'Total_Emission_Grazing_{M}']
        print(f"  ✓ 4.1 done — Total MUN: {df['Total_Emission_All_MUN'].iloc[0]:.2f}, VCRE: {df['Total_Emission_All_VCRE'].iloc[0]:.2f}")
        return df

# ══════════════════════════════════════════════════════════════════════════════
# MODULE 4.2 — Land Application (VCRE-based, RVO 2025 rules)
# ══════════════════════════════════════════════════════════════════════════════
#
# Key logic per your instruction:
#   "remaining TAN from total manure excretion is usually MORE than the
#    usage space → so N applied = usage space (the limit is binding)"
#
# Steps per farm:
#   1. Determine manure N limit (170/190/200 kg N/ha, based on derogation)
#   2. Determine total N usage norms per land parcel (Tabel 2)
#   3. Manure N available = N excretion (VCRE) after emission losses
#   4. Manure N applied = min(available, limit × ha)   [limit is binding]
#   5. Effective manure N = applied × working coefficient (Tabel 9)
#   6. Fertiliser N space = total N norm − effective manure N
# ══════════════════════════════════════════════════════════════════════════════


class LandApplicationCalculator:
    """
    Computes:
      - Total N usage space (kg N) based on land areas and crop norms
      - Manure N application limit (kg N) based on derogation/NV area
      - Available manure N (VCRE excretion minus NH3-N losses)
      - Applied manure N = min(available, manure_limit)
      - Effective manure N for N-space = applied * working_coefficient
      - Fertiliser N space = usage_space - effective_manure_N

    Expected columns (from earlier modules / input):
      Soil_Type, Region, Derogation, NV_Area, CropType, Ha_Crop, Ha_Grass, Ha_Mais,
      GD_* and slurry%_koe (for working coefficient), and emissions columns from Module 5.
    """

    # ── Tabel 2: N use norms (kg N/ha/yr) ─────────────────────────────────────
    TABEL_2 = {
        'grasland_beweiden':        {'klei': 345, 'zand_nwc': 250, 'zand_zuid': 250, 'loss': 250, 'veen': 265},
        'grasland_maaien':          {'klei': 385, 'zand_nwc': 320, 'zand_zuid': 320, 'loss': 320, 'veen': 300},
        'mais_derogatie':           {'klei': 160, 'zand_nwc': 140, 'zand_zuid': 112, 'loss': 112, 'veen': 150},
        'mais_geen_derogatie':      {'klei': 185, 'zand_nwc': 140, 'zand_zuid': 112, 'loss': 112, 'veen': 150},
        'consumptieaardappel_hoog': {'klei': 275, 'zand_nwc': 260, 'zand_zuid': 208, 'loss': 204, 'veen': 270},
        'consumptieaardappel_overig': {'klei': 250, 'zand_nwc': 235, 'zand_zuid': 188, 'loss': 184, 'veen': 245},
        'wintertarwe':              {'klei': 245, 'zand_nwc': 160, 'zand_zuid': 160, 'loss': 190, 'veen': 160},
        'zomertarwe':               {'klei': 150, 'zand_nwc': 140, 'zand_zuid': 140, 'loss': 140, 'veen': 140},
        'wintergerst':              {'klei': 140, 'zand_nwc': 140, 'zand_zuid': 140, 'loss': 140, 'veen': 140},
        'suikerbieten':             {'klei': 150, 'zand_nwc': 145, 'zand_zuid': 116, 'loss': 116, 'veen': 145},
        'zetmeelaardappelen':       {'klei': 240, 'zand_nwc': 230, 'zand_zuid': 184, 'loss': 184, 'veen': 230},
        'overig':                   {'klei': 200, 'zand_nwc': 185, 'zand_zuid': 148, 'loss': 148, 'veen': 190},
    }

    # ── "Tabel 9" working coefficients (apparent N effectiveness) ─────────────
    WC_SLURRY_GRAZING = 0.45
    WC_SLURRY_NOGRAZING = 0.60
    WC_SOLID_GRAZING = 0.45
    WC_SOLID_NOGRAZING = 0.60

    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()
        self._ensure_columns()

    def _ensure_columns(self):
        """Make sure required columns exist; coerce numerics."""
        # Normalize HA_Crop capitalization if needed
        if 'HA_Crop' in self.df.columns and 'Ha_Crop' not in self.df.columns:
            self.df.rename(columns={'HA_Crop': 'Ha_Crop'}, inplace=True)

        # Defaults
        if 'CropType' not in self.df.columns:
            self.df['CropType'] = 'overig'

        for c in ['Ha_Grass', 'Ha_Mais', 'Ha_Crop', 'NatureGL%']:
            self.df = ensure_numeric(self.df, c, 0.0)

        for c in ['GD_Limited_Koe', 'GD_Combi_Koe', 'GD_Unlimited_Koe', 'slurry%_koe', 'Derogation', 'NV_Area']:
            if c in ['Derogation', 'NV_Area']:
                if c not in self.df.columns:
                    self.df[c] = 'No'
            else:
                self.df = ensure_numeric(self.df, c, 0.0)

        # slurry% fraction handling
        if self.df['slurry%_koe'].max() > 1.0:
            self.df['slurry%_koe'] = (self.df['slurry%_koe'] / 100.0).clip(0.0, 1.0)
        else:
            self.df['slurry%_koe'] = self.df['slurry%_koe'].clip(0.0, 1.0)

        if 'Soil_Type' not in self.df.columns:
            self.df['Soil_Type'] = 'zand'
        if 'Region' not in self.df.columns:
            self.df['Region'] = 'Others'

    @staticmethod
    def _yes(x) -> bool:
        return str(x).strip().lower() in ('yes', 'ja', 'true', '1', 'y')

    @staticmethod
    def _soil_key(soil_type, region):
        s = str(soil_type).lower()
        r = str(region).lower()
        if 'klei' in s:
            return 'klei'
        if 'veen' in s:
            return 'veen'
        if any(t in s for t in ('loss', 'loes', 'löss')):
            return 'loss'
        if 'zand' in s:
            return 'zand_zuid' if ('zuid' in r or 'south' in r) else 'zand_nwc'
        return 'zand_nwc'

    def _crop_norm(self, crop_key, soil_key) -> float:
        ck = str(crop_key).strip().lower().replace('ï', 'i')
        entry = self.TABEL_2.get(ck, self.TABEL_2['overig'])
        return float(entry.get(soil_key, 185.0))

    def _grass_key(self, row) -> str:
        gd = float(row.get('GD_Limited_Koe', 0) or 0) + float(row.get('GD_Combi_Koe', 0) or 0) + float(row.get('GD_Unlimited_Koe', 0) or 0)
        return 'grasland_beweiden' if gd > 0 else 'grasland_maaien'

    def _maize_key(self, row) -> str:
        derog = self._yes(row.get('Derogation', 'No'))
        return 'mais_derogatie' if derog else 'mais_geen_derogatie'

    def _manure_limit_per_ha(self, row) -> float:
        derog = self._yes(row.get('Derogation', 'No'))
        nv = self._yes(row.get('NV_Area', 'No'))
        if derog:
            return 190.0 if nv else 200.0
        return 170.0

    def _is_grazing(self, row) -> bool:
        gd = float(row.get('GD_Limited_Koe', 0) or 0) + float(row.get('GD_Combi_Koe', 0) or 0) + float(row.get('GD_Unlimited_Koe', 0) or 0)
        return gd > 0

    def _working_coefficient(self, row) -> float:
        slurry_frac = float(row.get('slurry%_koe', 1.0) or 1.0)
        solid_frac = 1.0 - slurry_frac
        if self._is_grazing(row):
            return slurry_frac * self.WC_SLURRY_GRAZING + solid_frac * self.WC_SOLID_GRAZING
        return slurry_frac * self.WC_SLURRY_NOGRAZING + solid_frac * self.WC_SOLID_NOGRAZING

    def calculate_land_application(self) -> pd.DataFrame:
        print(f"\n{'='*70}\nMODULE 4.2: Land Application (VCRE-based)\n{'='*70}")
        df = self.df

        # Soil key & per-ha norms
        df['Soil_Key_5_1'] = [self._soil_key(s, r) for s, r in zip(df['Soil_Type'], df['Region'])]

        # Total land area for manure placement limit:
        # - Use grass + maize + crop. (If you need a different definition, adjust here.)
        df['Ha_Total_For_ManureLimit'] = (df['Ha_Grass'] + df['Ha_Mais'] + df['Ha_Crop']).clip(lower=0)

        # Manure placement limit (kg N/ha) and total (kg N/farm)
        df['Manure_Limit_kgN_per_ha'] = df.apply(self._manure_limit_per_ha, axis=1)
        df['Manure_Limit_Total_kgN'] = df['Manure_Limit_kgN_per_ha'] * df['Ha_Total_For_ManureLimit']

        # N usage norms ("gebruiksruimte") per land type
        # Grass norm depends on grazing
        grass_key = df.apply(self._grass_key, axis=1)
        maize_key = df.apply(self._maize_key, axis=1)

        df['Norm_Grass_kgN_per_ha'] = [self._crop_norm(k, sk) for k, sk in zip(grass_key, df['Soil_Key_5_1'])]
        df['Norm_Maize_kgN_per_ha'] = [self._crop_norm(k, sk) for k, sk in zip(maize_key, df['Soil_Key_5_1'])]
        df['Norm_Crop_kgN_per_ha'] = [self._crop_norm(k, sk) for k, sk in zip(df['CropType'], df['Soil_Key_5_1'])]

        df['UsageSpace_Grass_kgN'] = df['Norm_Grass_kgN_per_ha'] * df['Ha_Grass']
        df['UsageSpace_Maize_kgN'] = df['Norm_Maize_kgN_per_ha'] * df['Ha_Mais']
        df['UsageSpace_Crop_kgN'] = df['Norm_Crop_kgN_per_ha'] * df['Ha_Crop']
        df['UsageSpace_Total_kgN'] = df['UsageSpace_Grass_kgN'] + df['UsageSpace_Maize_kgN'] + df['UsageSpace_Crop_kgN']

        # Available manure N (VCRE) after NH3 emissions.
        # Use total VCRE excretion and subtract NH3-N losses (Stable+Storage+Grazing) in N units.
        for c in ['Total_N_Excretion_Farm_VCRE', 'Total_Emission_All_VCRE']:
            if c not in df.columns:
                df[c] = 0.0
            df = ensure_numeric(df, c, 0.0)

        df['ManureN_Available_After_Emissions_kgN'] = (df['Total_N_Excretion_Farm_VCRE'] - df['Total_Emission_All_VCRE']).clip(lower=0)

        # Applied manure N is capped by manure placement limit
        df['ManureN_Applied_kgN'] = np.minimum(df['ManureN_Available_After_Emissions_kgN'], df['Manure_Limit_Total_kgN'])

        # Working coefficient (effective fraction counting towards usage space)
        df['Working_Coefficient'] = df.apply(self._working_coefficient, axis=1)

        # Effective manure N for usage space
        df['ManureN_Effective_kgN'] = df['ManureN_Applied_kgN'] * df['Working_Coefficient']

        # Remaining fertiliser space
        df['FertiliserSpace_kgN'] = (df['UsageSpace_Total_kgN'] - df['ManureN_Effective_kgN']).clip(lower=0)

        # Optional: show which constraint binds
        df['BindingConstraint'] = np.where(
            df['ManureN_Available_After_Emissions_kgN'] >= df['Manure_Limit_Total_kgN'],
            'ManureLimit',
            'Availability'
        )

        print(f"  ✓ 4.2 done — Farm1 UsageSpace: {df['UsageSpace_Total_kgN'].iloc[0]:.1f} kg N; "
              f"Applied manure: {df['ManureN_Applied_kgN'].iloc[0]:.1f} kg N "
              f"({df['BindingConstraint'].iloc[0]})")
        return df

# ==============================================================================
# MODULE 4.3: Application Emission Calculation (Manure & Fertiliser)
# ==============================================================================

class ApplicationEmissionCalculator:
    """
    Calculates NH3 emissions from the application of Manure and Synthetic Fertilizers.
    Integrates results from stable/storage/grazing (calculated previously via EF & NEMA methods).
    """
    
    # --- EF Tables (from the provided screenshots) ---
    # We remove all spaces and lowercase everything to make matching immune to Excel typos
    
    # 1. Synthetic Fertiliser EFs (% of N applied)
    EF_FERTILISER = {
        '100%ammonium': 11.3,
        '100%nitraat': 0.0,
        'combinatievanammoniumetnitraat': 2.5, # Typo in excel matched here
        'combinatievanammoniumennitraat': 2.5, # Correct spelling fallback
        'ureum,gekorreld,zonderurease-remmer': 14.3,
        'ureum,gekorreld,meturease-remmer': 5.9,
        'voeibaarureumzonderurease-remmerofzuur': 7.5, # Typo 'voeibaar'
        'vloeibaarureumzonderurease-remmerofzuur': 7.5,
        'vloeibaarureummeturease-remmerofzuuer': 3.1,  # Typo 'zuuer'
        'vloeibaarureummeturease-remmerofzuur': 3.1,
        'vloeibaarureumtoegviainjectie': 1.5,          # Typo 'toeg'
        'vloeibaarureumtoegediendviainjectie': 1.5
    }

    # 2. Manure Application EFs on Grassland (% of TAN applied)
    EF_MANURE_GRASS = {
        'bovengronds': 68.0,
        'sleepvoet': 26.4,
        'sleepvoetverdund': 17.0,
        'sleufkouterverdund': 17.0,
        'sleufkouter': 21.7,
        'zodebemester': 17.0
    }

    # 3. Manure Application EFs on Cropland (Bouwland) (% of TAN applied)
    EF_MANURE_ARABLE = {
        'bovengronds': 69.0,
        'ineenwerkgangonderwerken': 22.0,
        'sleepvoet': 36.0,
        'diepeinjectie': 2.0,
        'ondiepeinjectie': 24.0
    }

    def __init__(self, df):
        self.df = df.copy()

    def _clean_string(self, s):
        """Helper to lowercase and strip ALL spaces for robust dictionary matching"""
        return str(s).lower().replace(' ', '').strip()

    def run_application_emission(self):
        print(f"\n{'='*70}\nMODULE 4.3: Application Emissions\n{'='*70}")
        
        # ---------------------------------------------------------
        # 1. Lookup Emission Factors (EF)
        # ---------------------------------------------------------
        
        # Synthetic Fertilizer EF
        fert_forms = self.df.get('Fertiliser_Form', pd.Series('100%ammonium', index=self.df.index))
        cleaned_fert_forms = fert_forms.apply(self._clean_string)
        self.df['EF_Fertiliser_%'] = cleaned_fert_forms.map(self.EF_FERTILISER).fillna(0.0)

        # Manure Application EF (Grassland)
        tech_grass = self.df.get('AM_Grassland', pd.Series('zodebemester', index=self.df.index))
        self.df['EF_Manure_Grass_%'] = tech_grass.apply(self._clean_string).map(self.EF_MANURE_GRASS).fillna(17.0)
        
        # Manure Application EF (Arable/Cropland)
        tech_arable = self.df.get('AM_Cropland', pd.Series('ondiepeinjectie', index=self.df.index))
        self.df['EF_Manure_Arable_%'] = tech_arable.apply(self._clean_string).map(self.EF_MANURE_ARABLE).fillna(24.0)

        # ---------------------------------------------------------
        # 2. Calculate TAN Applied & Manure Emissions (Average Method)
        # ---------------------------------------------------------
        
        # --- A. Retrieve correct column data from previous modules ---
        tan_correct_mun = self.df['Total_Corrected_TAN_MUN']
        prev_emission_mun = self.df['Total_Emission_All_MUN']
        
        # Handle potential column name variations from Module 1.1 safely
        if 'Total_Nitrogen_Excretion_kg' in self.df.columns:
            n_excr_mun = self.df['Total_Nitrogen_Excretion_kg']
        else:
            n_excr_mun = self.df.get('Total_Nitrogen_Excretion_MUN', 0.0) 
        
        tan_correct_vcre = self.df['Total_Corrected_TAN_VCRE']
        prev_emission_vcre = self.df['Total_Emission_All_VCRE']
        n_excr_vcre = self.df['Total_N_Excretion_Farm_VCRE'] 
        
        # Nitrogen input for the application stage (from Module 4.2)
        n_applied_manure = self.df['ManureN_Applied_kgN']
        n_applied_fert = self.df['FertiliserSpace_kgN']
        
        # --- B. Calculate Net TAN (deduct previous NH3 emissions, multiply by 14/17 to convert to N) ---
        net_tan_mun = (tan_correct_mun - (prev_emission_mun * (14.0 / 17.0))).clip(lower=0)
        net_tan_vcre = (tan_correct_vcre - (prev_emission_vcre * (14.0 / 17.0))).clip(lower=0)
        
        # --- C. Calculate averages (Numerator & Denominator) ---
        avg_net_tan = (net_tan_mun + net_tan_vcre) / 2.0
        self.df['Avg_Net_TAN_Excreted'] = avg_net_tan
        
        avg_n_excr = (n_excr_mun + n_excr_vcre) / 2.0
        avg_n_excr_safe = np.where(avg_n_excr > 0, avg_n_excr, 1.0) # Prevent division by zero
        
        # --- D. Calculate unified TAN % ---
        tan_pct = avg_net_tan / avg_n_excr_safe
        self.df['Avg_TAN_Pct'] = tan_pct.clip(0, 1) # Cap between 0 and 100%
        
        # ---------------------------------------------------------
        # 3. Calculate Manure Application Emissions (Land-Specific)
        # ---------------------------------------------------------
        
        # Get land areas (assuming arable land = Maize + Other Crops)
        ha_grass = self.df.get('Ha_Grass', 0.0)
        ha_mais = self.df.get('Ha_Mais', 0.0)
        ha_crop = self.df.get('Ha_Crop', 0.0)
        
        total_ha = ha_grass + ha_mais + ha_crop
        total_ha_safe = np.where(total_ha > 0, total_ha, 1.0) 
        
        # Calculate the fraction of manure going to each land type (proportional to area)
        frac_grass = ha_grass / total_ha_safe
        frac_arable = (ha_mais + ha_crop) / total_ha_safe
        
        # Distribute Applied Manure N per land type
        applied_n_grass = n_applied_manure * frac_grass
        applied_n_arable = n_applied_manure * frac_arable
        
        # Calculate TAN Applied per land type
        tan_applied_grass = applied_n_grass * self.df['Avg_TAN_Pct']
        tan_applied_arable = applied_n_arable * self.df['Avg_TAN_Pct']
        
        self.df['TAN_Applied_Manure_Grass'] = tan_applied_grass
        self.df['TAN_Applied_Manure_Arable'] = tan_applied_arable
        self.df['TAN_Applied_Manure_Total'] = tan_applied_grass + tan_applied_arable
        
        # Calculate emissions using land-specific EFs (convert NH3-N back to NH3 gas via 17/14)
        emission_grass = (tan_applied_grass * self.df['EF_Manure_Grass_%'] / 100.0) * (17.0 / 14.0)
        emission_arable = (tan_applied_arable * self.df['EF_Manure_Arable_%'] / 100.0) * (17.0 / 14.0)
        
        app_emission_manure = emission_grass + emission_arable
        
        self.df['Emission_ManureApp_Grass'] = emission_grass
        self.df['Emission_ManureApp_Arable'] = emission_arable
        self.df['Emission_ManureApp_Total'] = app_emission_manure
        
        # ---------------------------------------------------------
        # 4. Calculate Fertiliser Emission
        # ---------------------------------------------------------

        app_emission_fert = (n_applied_fert * self.df['EF_Fertiliser_%'] / 100.0) * (17.0 / 14.0)
        self.df['Emission_FertiliserApp'] = app_emission_fert

        # ---------------------------------------------------------
        # 5. Final Total Emissions & GVE Normalization
        # ---------------------------------------------------------
        # Calculate final total emissions separately for both MUN and VCRE systems
        self.df['Total_Farm_NH3_Emission_MUN'] = prev_emission_mun + app_emission_manure + app_emission_fert
        self.df['Total_Farm_NH3_Emission_VCRE'] = prev_emission_vcre + app_emission_manure + app_emission_fert
            
        # 🌟 1. Calculate GVE (Grootvee-eenheid) 🌟
        # Standard Dutch (CBS/RVO) GVE factors: Cow=1.0, Pink=0.5, Kalf=0.25
        self.df['Total_GVE_Farm'] = (self.df['Nr_koe'] * 1.0) + (self.df['Nr_pink'] * 0.50) + (self.df['Nr_kalf'] * 0.25)
        
        # Prevent division by zero
        safe_gve = np.where(self.df['Total_GVE_Farm'] > 0, self.df['Total_GVE_Farm'], 1.0)
        
        # 🌟 2. Normalize TOTAL Emissions per GVE 🌟
        self.df['NH3_Emission_per_GVE_MUN'] = self.df['Total_Farm_NH3_Emission_MUN'] / safe_gve
        self.df['NH3_Emission_per_GVE_VCRE'] = self.df['Total_Farm_NH3_Emission_VCRE'] / safe_gve

        # 🌟 3. NEW: Normalize STABLE and FERTILISER Emissions per GVE 🌟
        self.df['Stable_Emission_per_GVE_MUN'] = self.df['Total_Emission_Stable_MUN'] / safe_gve
        self.df['Stable_Emission_per_GVE_VCRE'] = self.df['Total_Emission_Stable_VCRE'] / safe_gve
        self.df['Fertiliser_Emission_per_GVE'] = self.df['Emission_FertiliserApp'] / safe_gve

        print(f"  ✓ 4.3 Application Emissions Calculated")
        print(f"      Farm GVE: {self.df['Total_GVE_Farm'].iloc[0]:.1f}, Total NH3/GVE (VCRE): {self.df['NH3_Emission_per_GVE_VCRE'].iloc[0]:.1f} kg")
        return self.df

def run_pipeline(input_file, sheet_name='Main input'):
    """
    Runs the entire REMAS pipeline on the provided input file.
    input_file can be a file path OR a Streamlit uploaded file object.
    """
    # ── 1.1 Load + manure totals
    m11 = FarmManureCalculator(input_file, sheet_name)
    df = m11.load_and_clean()
    m11.df = df
    df = m11.calculate_totals()

    # ── 1.2 MUN partitioning
    df = NitrogenPartitioningCalculator(df).calculate_partitioning()

    # ── 2.1 VEM requirements
    df = VEMRequirementCalculator(df).calculate_requirements()

    # ── 2.2 VEM allocation
    df = VEMAllocationCalculator(df).run_allocation()

    # ── 2.3 N intake
    df = NitrogenIntakeCalculator(df).run_nutrient_calculation()

    # ── 2.4 VCRE excretion
    df = NitrogenExcretionCalculatorVCRE(df).calculate_excretion()

    # ── 2.5 VCRE partitioning
    df = NitrogenPartitioningVCRE(df).calculate_partitioning()

    # ── 3.1 mineralization
    df = MineralizationCalculator(df).calculate_mineralization()

    # ── 3.2 corrected TAN
    df = CorrectedTANCalculator(df).calculate_corrected_tan()

    # ── 4.1 emissions during stable, grazing, storage
    df = EmissionCalculator(df).calculate_emissions()

    # ── 4.2 land application (VCRE)
    df = LandApplicationCalculator(df).calculate_land_application()
    
    # ── 4.3 Application Emission
    df = ApplicationEmissionCalculator(df).run_application_emission()

    # Return the final dataframe instead of saving to a local hard drive
    return df
    