import pandas as pd
from app import app
from models import db, FormTemplate

def import_from_excel(file_path):
    try:
        # قراءة ملف Excel
        df = pd.read_excel(file_path)
        
        with app.app_context():
            # مسح البيانات القديمة
            FormTemplate.query.delete()
            
            # استيراد البيانات الجديدة
            for index, row in df.iterrows():
                template = FormTemplate(
                    field=row['المجال'],
                    standard_number=row['المعيار'],
                    description=row['التفاصيل'],
                    weight=row['الوزن'],
                    rating_0=row['0'],
                    rating_1=row['1'],
                    rating_2=row['2'],
                    rating_3=row['3'],
                    rating_4=row['4']
                )
                db.session.add(template)
            
            db.session.commit()
            print("✅ تم استيراد النموذج من Excel بنجاح")
            
    except Exception as e:
        print(f"❌ خطأ في الاستيراد: {e}")

if __name__ == "__main__":
    import_from_excel('visit_form_template.xlsx')