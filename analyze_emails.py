import json
import re
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

from datetime import datetime
try:
    import google.generativeai as genai
    HAS_GENAI = True
except ImportError:
    HAS_GENAI = False

def analyze_with_llm(sender, subject, body):
    """Analyze email using Gemini 1.5 Pro to determine client and matter."""
    api_key = os.environ.get("GEMINI_API_KEY")
    if not HAS_GENAI or not api_key:
        return None

    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-pro-latest')
        
        # Truncate body to prevent excessive token usage while keeping context
        clean_body = body[:4000]
        
        prompt = f"""
        You are a legal assistant categorizing emails. Analyze this email metadata and content to identify the Client and Matter/Project.
        
        Sender: {sender}
        Subject: {subject}
        Body Start: {clean_body}
        
        Task:
        1. Identify the Client (Company/Organization name). If internal or personal, specify.
        2. Identify the Matter or Project (e.g., "Contract Review", "GDPR Compliance", "Employment Dispute").
        3. Rate confidence from 1-5 (5 is certain).
        
        Return ONLY a JSON object with this structure:
        {{
            "proposedClient": "string",
            "proposedMatter": "string",
            "confidence": integer,
            "reasoning": "string"
        }}
        """
        
        response = model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
        return json.loads(response.text)
    except Exception as e:
        print(f"LLM analysis failed: {e}")
        return None

def analyze_email(email):
    """Analyze an email and add client/matter/confidence properties."""
    
    subject = email.get("Subject", "").lower()
    body = email.get("Body", "").lower()
    sender_email = email.get("SenderEmailAddress", "").lower()
    sender_name = email.get("SenderName", "")
    
    # Try LLM Analysis first
    llm_result = analyze_with_llm(f"{sender_name} <{sender_email}>", email.get("Subject", ""), email.get("Body", ""))
    if llm_result and llm_result.get("confidence", 0) >= 3:
        email["proposedClient"] = llm_result.get("proposedClient", "None")
        email["proposedMatter"] = llm_result.get("proposedMatter", "N/A")
        email["confidenceRating"] = llm_result.get("confidence", 0)
        email["aiReasoning"] = llm_result.get("reasoning", "")
        return email

    # Extract domain from sender
    domain_match = re.search(r'@([a-z0-9.-]+)', sender_email)
    sender_domain = domain_match.group(1) if domain_match else ""
    
    proposed_client = "None"
    proposed_matter = "N/A"
    confidence = 1
    
    # Check if internal Harper James email
    is_exchange = "/O=EXCHANGELABS" in sender_email or "/O=HARPERJAMES" in sender_email
    is_internal_domain = "harperjames.co.uk" in sender_email or sender_domain == "harperjames.co.uk"
    is_internal = is_internal_domain or is_exchange
    
    # Client detection patterns
    client_patterns = {
        "Policy in Practice": ["policy in practice", "pip", "west northampton"],
        "South Pole": ["south pole", "luumo"],
        "AI Health Tech Ltd": ["ai health tech", "abdul kamali", "nicholas dacre"],
        "Oriel Services": ["oriel services", "oriel"],
        "Prinsix": ["prinsix", "kcom"],
        "Integrated Doorset Solutions": ["integrated doorset", "mark higgs"],
        "Streets Heaver": ["streets heaver"],
        "CiteAb": ["citeab"],
        "Biorelate": ["biorelate"],
        "Avallo": ["avallo"],
        "Blackbird": ["blackbird plc"],
        "CF Psychology": ["cf psychology"],
        "Extend": ["extend"],
    }
    
    # Matter detection patterns
    matter_keywords = {
        "dpa": "Data Processing Agreement",
        "data processing agreement": "Data Processing Agreement",
        "saas agreement": "SaaS Agreement",
        "software agreement": "Software Agreement",
        "contract": "Contract Review",
        "foi": "FOI Request",
        "freedom of information": "FOI Request",
        "gdpr": "GDPR Compliance",
        "sow": "Statement of Work",
        "terms and conditions": "Terms & Conditions",
        "nda": "Non-Disclosure Agreement",
        "employment": "Employment Matter",
        "dismissal": "Employment - Dismissal",
    }
    
    # Check for client mentions
    text_to_check = subject + " " + body
    for client, keywords in client_patterns.items():
        if any(keyword in text_to_check for keyword in keywords):
            proposed_client = client
            confidence = 3
            break
    
    # Check for matter type
    for keyword, matter_type in matter_keywords.items():
        if keyword in text_to_check:
            proposed_matter = matter_type
            if proposed_client != "None":
                confidence = 4
            break
    
    # Specific patterns for higher confidence
    if is_internal:
        if proposed_client != "None":
            # Internal discussion about client work
            proposed_client = f"{proposed_client} (Internal)"
            confidence = max(3, confidence)
        else:
            # Pure internal email
            proposed_client = "Harper James (Internal)"
            
            # Check for specific internal topics
            if any(word in text_to_check for word in ["meeting", "work stack", "team", "catch up", "huddle"]):
                proposed_matter = "Team/Meeting"
                confidence = 5
            elif any(word in text_to_check for word in ["system", "bug", "error", "technical", "mcp", "core"]):
                proposed_matter = "IT/Technical"
                confidence = 5
            elif any(word in text_to_check for word in ["marketing", "newsletter", "brand", "webinar", "press", "media", "social media"]):
                proposed_matter = "Marketing"
                confidence = 5
            elif any(word in text_to_check for word in ["holiday", "absence", "annual leave", "sickness", "hr", "benefits", "appraisal", "promotion", "training", "competence"]):
                proposed_matter = "HR & Operations"
                confidence = 5
            elif any(word in text_to_check for word in ["utilisation", "billing", "wip", "finance", "invoice", "rates"]):
                proposed_matter = "Finance & Management"
                confidence = 5
            elif "extend client list" in text_to_check:
                proposed_matter = "Business Development"
                confidence = 5
            else:
                proposed_matter = "General Internal"
                confidence = 3
    
    # If clear client and matter identified with specific details
    if proposed_client != "None" and proposed_client != "Internal":
        if proposed_matter != "N/A":
            # Check for urgency or specific project names
            if any(word in text_to_check for word in ["urgent", "asap", "deadline"]):
                confidence = 5
            elif re.search(r'\b(project|agreement|contract)\s+\w+', text_to_check):
                confidence = 5
    
    # Add properties to email
    email["proposedClient"] = proposed_client
    email["proposedMatter"] = proposed_matter
    email["confidenceRating"] = confidence
    
    return email

def main():
    # Read the JSON file
    input_file = r"c:\Users\DavidSant\effi-mail\Untitled-1.json"
    
    print(f"Reading {input_file}...")
    with open(input_file, 'r', encoding='utf-8') as f:
        emails = json.load(f)
    
    print(f"Found {len(emails)} emails to analyze...")
    
    # Analyze each email
    for i, email in enumerate(emails):
        if i % 10 == 0:
            print(f"Processing email {i+1}/{len(emails)}...")
        emails[i] = analyze_email(email)
    
    # Write back to file
    output_file = input_file
    print(f"\nWriting results to {output_file}...")
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(emails, f, indent=4, ensure_ascii=False)
    
    # Print summary
    print("\n=== Analysis Complete ===")
    client_counts = {}
    for email in emails:
        client = email.get("proposedClient", "Unknown")
        client_counts[client] = client_counts.get(client, 0) + 1
    
    print("\nClient Distribution:")
    for client, count in sorted(client_counts.items(), key=lambda x: x[1], reverse=True):
        print(f"  {client}: {count}")
    
    confidence_counts = {}
    for email in emails:
        conf = email.get("confidenceRating", 0)
        confidence_counts[conf] = confidence_counts.get(conf, 0) + 1
    
    print("\nConfidence Distribution:")
    for conf in sorted(confidence_counts.keys(), reverse=True):
        print(f"  Level {conf}: {confidence_counts[conf]}")

if __name__ == "__main__":
    main()
