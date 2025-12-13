"""
Add explanatory comments to the PDX DSA document for client review.
Each paragraph gets a comment explaining its purpose.
"""

import win32com.client
import os

def add_comments_to_dsa(docx_path):
    """Add explanatory comments to the PDX DSA document"""
    
    # Open Word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True  # Show Word so client can see the result
    
    try:
        doc = word.Documents.Open(docx_path)
        
        # Dictionary mapping paragraph text patterns to comments
        comments = {
            # Cover/Parties
            "DATA SHARING AGREEMENT": "This is the title of the agreement - a formal contract for sharing credit data with the CRA.",
            "PERCAYSO LIMITED": "This identifies Percayso as party (1), acting as agent for all the lenders. This agency structure means one agreement covers all your lender clients.",
            "[CRA NAME]": "Party (2) is the Credit Reference Agency. The specific CRA name will be inserted when the agreement is finalised.",
            
            # Background/Recitals
            "Percayso operates a data sharing platform": "Background (A): Sets the scene by explaining what PDX does - it's your platform that facilitates credit data sharing.",
            "Each Participating Lender has entered into": "Background (B): Explains that each lender has already signed up with Percayso and authorised you to act as their agent.",
            "The CRA wishes to receive Credit Information": "Background (C): Confirms the CRA's interest in receiving the data under the Regulations.",
            "Percayso enters into this Agreement as agent": "Background (D): Key clause - makes clear that Percayso is acting as agent, and each lender is bound as if they signed directly.",
            
            # Clause 1: Definitions
            "DEFINITIONS AND INTERPRETATION": "CLAUSE 1: Standard definitions section. Defines key terms used throughout the agreement.",
            "In this Agreement the following words": "Introduction to the definitions list.",
            '"Agreement" means': "Defines what 'Agreement' covers - the whole document including schedules.",
            '"Applicable Laws" means': "Covers all relevant legislation - importantly includes the Regulations, DPA 2018, and UK GDPR.",
            '"Confidential Information" means': "Broad definition of confidential information for the confidentiality clause.",
            '"CRA" means': "Identifies this is a designated CRA under the SME Credit Information Regulations.",
            '"Credit Information" means': "The actual data being shared - defined by reference to Schedule 2.",
            '"Credit Provider" means': "Defines credit providers (e.g. trade credit) as distinct from Finance Providers.",
            '"Effective Date" means': "When the agreement starts - the date of signing.",
            '"Finance Provider" shall be interpreted': "Cross-refers to the statutory definition in the Regulations (includes banks).",
            '"Group Company" means': "Standard group company definition for any provisions dealing with affiliates.",
            '"Intellectual Property Rights" means': "Comprehensive IP definition for the IP ownership clause.",
            '"Lender Agreement" means': "Your agreement with each lender - this is where they appoint you as agent.",
            '"Participating Lender" means': "Defines who is covered - those in Schedule 1 who haven't left.",
            '"PDX Platform" means': "Defines your platform's role - validation, transformation and submission.",
            '"Percayso" means': "Makes clear Percayso acts as agent, not principal.",
            '"Principles of Reciprocity" means': "Industry standard principles for credit data sharing between CRAs and lenders.",
            '"Regulations" means': "The SME Credit Information Regulations 2015 - the statutory framework for this agreement.",
            '"UK GDPR" means': "Data protection law definition - the retained EU GDPR.",
            "In this Agreement:": "Introduction to interpretation rules.",
            "any reference to a statutory provision": "Standard boilerplate - legislation references include future amendments.",
            "references to clauses and schedules": "Confirms internal cross-references.",
            "the singular includes the plural": "Standard drafting convention.",
            "headings are for ease of reference": "Headings don't affect interpretation.",
            "where any matter is to be agreed": "Agreements must be in writing.",
            'wherever the words "including"': "Makes 'including' non-exhaustive - important for flexibility.",
            
            # Clause 2: Agency
            "AGENCY": "CLAUSE 2: The core agency provisions. This is the key structural innovation - Percayso acts as agent for all lenders.",
            "Percayso enters into this Agreement as agent for and on behalf": "Establishes the agency relationship - each lender is bound as if they signed directly.",
            "Percayso warrants that it has been duly authorised": "Percayso's warranty that it has authority from each lender (via the Lender Agreement).",
            "Schedule 1 sets out the current list": "Lenders are listed in Schedule 1 and can be added/removed.",
            "To add a Participating Lender": "Process for adding new lenders - 10 business days' notice to CRA.",
            "To remove a Participating Lender": "Process for removing lenders - 30 days' notice.",
            "The CRA may request that Percayso remove": "CRA's right to request removal of a lender in material breach.",
            "For the avoidance of doubt, Percayso is not itself": "Clarifies Percayso's role - agent and service provider, not a data supplier.",
            
            # Clause 3: Supply of Credit Information
            "SUPPLY OF CREDIT INFORMATION": "CLAUSE 3: The mechanics of data supply - what, when, and how.",
            "Subject to paragraph 6(2) of the Regulations": "Data sharing obligation, subject to the 12-month window in the Regulations. Monthly updates thereafter.",
            "The Credit Information shall be provided in a format": "Format to be agreed; Percayso handles validation and transformation.",
            "Where testing of the Credit Information is required": "Testing provisions - CRA can't use test data in live systems.",
            "Each Participating Lender shall use reasonable endeavours": "Accuracy obligation sits with lenders; Percayso validates but doesn't warrant accuracy.",
            "If a Participating Lender becomes aware": "Correction process if inaccurate data is supplied.",
            
            # Clause 4: Licence
            "LICENCE": "CLAUSE 4: What the CRA can do with the data - a non-exclusive, royalty-free licence.",
            "Each Participating Lender grants to the CRA": "Introduction to the permitted uses list.",
            "to develop and sell products and services": "Core use - CRA can build credit products for identity verification, fraud prevention, creditworthiness assessment etc.",
            "to develop and sell scores to Trade Credit Providers": "Specific permission for trade credit scoring products.",
            "for the purposes necessary for the CRA to comply": "CRA's regulatory compliance use.",
            "to provide a copy of the Credit Information": "Businesses can request their own data from the CRA.",
            "for such other purposes as are consistent": "Flexibility for other uses consistent with Principles of Reciprocity.",
            "The licence to use the Credit Information shall expire": "Licence ends when lender leaves, but CRA can retain for regulatory compliance.",
            
            # Clause 5: Flow-Down Warranties
            "FLOW-DOWN WARRANTIES": "CLAUSE 5: What each lender has warranted to Percayso (passed through to CRA).",
            "Percayso warrants that, under each Lender Agreement": "Percayso confirms what lenders have already promised in their Lender Agreements.",
            "it is a Finance Provider or Credit Provider": "Lender confirms eligibility under Regulations.",
            "each customer in respect of whom": "Lender confirms customers have consented to data sharing.",
            "it will use reasonable endeavours to ensure": "Lender's accuracy warranty.",
            "it will notify Percayso promptly if it becomes aware": "Lender's correction obligation.",
            "it will comply with all Applicable Laws": "Lender's legal compliance warranty.",
            "it has appointed Percayso as its agent": "Lender confirms agent appointment.",
            "Percayso shall ensure that each Lender Agreement": "Lender Agreements must be at least as strong as these warranties.",
            "On reasonable request, Percayso shall provide": "CRA can review template Lender Agreement.",
            "At the CRA's request, Percayso shall use reasonable": "CRA can request sample customer notifications.",
            
            # Clause 6: Compliance and Data Protection
            "COMPLIANCE AND DATA PROTECTION": "CLAUSE 6: Legal compliance and GDPR provisions.",
            "Each party undertakes to the other": "Mutual compliance obligation.",
            "To the extent any Credit Information is personal data": "Data controller status - each lender controls its data; CRA becomes controller on receipt.",
            "Percayso acts as a data processor": "Clarifies Percayso's processor role for PDX Platform services.",
            "The CRA warrants that it shall take appropriate": "CRA's security obligations (GDPR Article 32).",
            "The CRA shall permit Percayso": "Annual audit right for Percayso on behalf of lenders.",
            "Each party agrees to promptly provide any information": "Cooperation on data subject requests and complaints.",
            
            # Clause 7: Liability
            "LIABILITY": "CLAUSE 7: Liability provisions - who is responsible for what.",
            "Neither party shall be liable for indirect": "Standard exclusion of indirect/consequential loss.",
            "Nothing in this Agreement shall operate to exclude": "Carve-outs for death/personal injury, fraud etc. (legally required).",
            "Percayso shall not be liable to the CRA": "Percayso not liable for lender breaches (except own negligence).",
            "Each Participating Lender shall be liable": "Each lender responsible for its own breaches.",
            
            # Clause 8: Term and Termination
            "TERM AND TERMINATION": "CLAUSE 8: Duration and how the agreement can end.",
            "This Agreement shall begin on the Effective Date": "Agreement runs until terminated.",
            "Either party may terminate this Agreement at any time": "90 days' notice termination right.",
            "the CRA's designation as a credit reference agency": "Immediate termination if CRA loses designated status.",
            "Percayso ceases to operate the PDX Platform": "Immediate termination if PDX shuts down.",
            "the other party commits a material breach": "Termination for material breach (immediate if not capable of remedy).",
            "which is not remedied within 28 days": "28 days to cure a remediable breach.",
            "the other party has passed a resolution": "Termination on insolvency.",
            "Termination of a Participating Lender's participation": "Individual lender leaving doesn't affect the main agreement.",
            "Upon termination of this Agreement, the CRA may retain": "Data retention/deletion on termination.",
            "Termination of this Agreement shall not affect": "Survival of accrued rights.",
            
            # Clause 9: IP
            "INTELLECTUAL PROPERTY RIGHTS": "CLAUSE 9: IP ownership - everyone keeps what they came in with.",
            "Any Intellectual Property Rights in the Credit Information": "Lenders own their data.",
            "All Intellectual Property Rights in the databases": "CRA owns its databases and derived products.",
            "All Intellectual Property Rights in the PDX Platform": "Percayso owns PDX.",
            
            # Clause 10: Confidentiality
            "CONFIDENTIALITY": "CLAUSE 10: Confidentiality obligations.",
            "Each party shall keep the Confidential Information": "Mutual confidentiality obligation.",
            "Each party may disclose Confidential Information to employees": "Permitted disclosures to staff/advisers.",
            "is or comes within the public domain": "Standard exceptions - public domain information.",
            "was in the recipient's possession": "Exception - already known information.",
            "is lawfully received from a third party": "Exception - third party information.",
            "is independently developed": "Exception - independent development.",
            "is required to be disclosed by law": "Exception - legal compulsion.",
            "The CRA shall not identify any Participating Lender": "CRA cannot reveal which lender provided specific data.",
            "This Clause 10 shall survive": "Confidentiality continues after termination.",
            
            # Clause 11: General
            "GENERAL": "CLAUSE 11: Boilerplate provisions.",
            "Any notices under this Agreement": "How to serve notices - writing, to registered office.",
            "Neither party may assign": "Assignment needs consent.",
            "Except as expressly provided, nothing in this Agreement shall create rights": "No third party rights (Contracts (Rights of Third Parties) Act 1999 exclusion).",
            "This Agreement sets out all the terms": "Entire agreement clause.",
            "Neither party will be liable for any delay": "Force majeure provision.",
            "This Agreement shall be governed by the laws of England": "English law and courts.",
            
            # Schedules
            "SCHEDULE 1": "List of participating lenders covered by this agreement.",
            "PARTICIPATING LENDERS": "The lenders who have signed up and authorised Percayso to act for them.",
            "SCHEDULE 2": "Details of the credit information to be shared.",
            "CREDIT INFORMATION": "Specifies exactly what data is included - follows the Regulations' requirements.",
            "Information relating to a loan": "Loan data fields as required by the Regulations.",
            "Information relating to a credit card": "Credit card data fields as required by the Regulations.",
            "Information relating to a current account": "Current account data fields as required by the Regulations.",
            "Where any of the information described above": "Business identification data that must accompany credit information.",
        }
        
        # Track added comments to avoid duplicates
        added_comments = set()
        
        # Iterate through paragraphs and add comments
        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            
            if not text or len(text) < 5:
                continue
            
            # Find matching comment
            for pattern, comment in comments.items():
                if pattern in text and pattern not in added_comments:
                    # Add comment to this paragraph
                    try:
                        doc.Comments.Add(para.Range, comment)
                        added_comments.add(pattern)
                        print(f"Added comment for: {pattern[:40]}...")
                    except Exception as e:
                        print(f"Could not add comment for '{pattern[:30]}': {e}")
                    break
        
        # Save the document
        doc.Save()
        print(f"\nDocument saved with {len(added_comments)} comments added.")
        print("The document is now open in Word for your review.")
        
    except Exception as e:
        print(f"Error: {e}")
        raise
    
    # Don't close Word - leave it open for the user to review
    # word.Quit()

if __name__ == '__main__':
    docx_path = r'C:\Users\DavidSant\OneDrive - Harper James Solicitors\Documents\Client Matters\Percayso\PDX DSA - Agency Model.docx'
    add_comments_to_dsa(docx_path)
    print("\nDone!")
