#!/usr/bin/python3.6

from fpdf import FPDF

SRC = 'source_files/'
CDS = 'bar_codes/'
LTTRS = 'letters/'

# FORMATS

FONT = 'helvetica'
F_SIZE = 12

PAGE_FMT = {'orientation': 'P',
            'unit': 'mm',
            'format': 'A4'}

IMAGE_ORDS = {'x': 10, 
              'y': 8, 
              'w': 100}

def make_pdf(PAGE_FMT):
    '''
    returns an instance of the FPDF class
    '''

    try:
        pdf = FPDF(**PAGE_FMT)
        pdf.add_page()

        return pdf
    except:
        raise Exception('ERROR: could not create pdf')

def add_txt(a_pdf, txt_str):

    try:
        a_pdf.set_font(FONT, size=F_SIZE)
        a_pdf.cell(200, 10, txt=txt_str, ln=1)
        a_pdf.ln(0.15)
                                            
    except:
        raise Exception(f'ERROR: could not add {txt_str}')

def add_image(a_pdf, image_path, ords=IMAGE_ORDS):
    '''
    a_pdf: a FPDF object initialized by make_pdf()
    image_path: path and file name of image to insert
    ords: dictionary of (x, y, width) values that are unpacked into the FPDF image
    method
    '''
    try:    
        a_pdf.image(image_path, **ords)
    except:
        raise Exception(f'ERROR: could not add image {image_path}')

def save_pdf(a_pdf, f_name):
    try:
        a_pdf.output(f_name)
    except:
        raise Exception(f'ERROR: could not save {f_name}')

def write_letter(a_image, applicant, app_date, app_email, services, location,
                 pu_date):
    '''
    wraps around the add_text() and add_image() function to pass in
    standard text for the letter
    and then calls the save_pdf() function naming the file after
    the applicant
    '''
    lttr_txt = [f'Dear {applicant}',
                'This letter is a confirmation of your Christmas Hamper Registration that you completed online',
                f'via the Christmas Bureau Website on {app_date} using {app_email}',
                f'You have registered for a package that will be a {services}',
                f'You will be able to pickup your {services} at {location}',
                f'on {pu_date}',
                'Please make sure you bring this letter with you when you pickup.',
                'We will need your account number to confirm your pickup', 
                'If you have any questions you can reach us during regular business hours at 555-555-5555',
                'or by email at info@christmashampers.ca'
               ]
    a_pdf = make_pdf(PAGE_FMT)
    add_image(a_pdf, f'{SRC}ChristmasBureauLogo.png', ords={})
    a_pdf.ln(35)
    for ln in lttr_txt:
        add_txt(a_pdf, ln)
    a_pdf.ln(40)
    # add text lable for bar code
    a_pdf.cell(200, 10, txt='Account Number',ln=1, align="C")
    # add cell to push image to middle of page
    a_pdf.cell(65, 10, align="C")
    add_image(a_pdf, a_image, ords={}) # add barcode image
    save_pdf(a_pdf, f'{LTTRS}{applicant}.pdf')

if __name__ == "__main__":
    # test run
    test_applicant = {'a_image': f'{CDS}118414.png' ,
                      'applicant': 'Mr. J F Smith', 
                      'app_date': 'Nov 14, 2020', 
                      'app_email': 'fake.email@fmail.com', 
                      'services': 'Gift Card', 
                      'location': 'North Community Centre',
                      'pu_date':'Dec 3 2020 at 2:00pm'}
    write_letter(**test_applicant)
