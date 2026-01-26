"""
Admin for managing the connection to the Forums backend service.
"""

from django import forms
from django.contrib import admin, messages
from django.contrib.auth.models import User
from django.contrib.admin.widgets import FilteredSelectMultiple
from django.db import models
from django.shortcuts import redirect

from openedx.core.djangoapps.content.course_overviews.models import CourseOverview
from openedx.core.djangoapps.enrollments.api import add_enrollment
from openedx.core.djangoapps.enrollments.data import CourseEnrollmentExistsError

from .models import ForumsConfig

admin.site.register(ForumsConfig)


class BulkEnrollment(models.Model):
    """Dummy model for admin interface"""
    class Meta:
        managed = False
        verbose_name = 'Bulk Enrollment'
        verbose_name_plural = 'Bulk Enrollment'
        app_label = 'django_comment_common'


class BulkEnrollmentForm(forms.Form):
    emails = forms.CharField(
        widget=forms.Textarea(attrs={'rows': 4, 'cols': 80}),
        help_text='Comma-separated: example@web.com, example2@web.com'
    )
    courses = forms.MultipleChoiceField(
        widget=forms.SelectMultiple(attrs={'size': 15, 'style': 'width: 500px;'}),
        help_text='Hold Ctrl/Cmd to select multiple courses'
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        course_choices = [
            (str(course.id), f"{course.display_name} ({course.id})")
            for course in CourseOverview.objects.all().order_by('display_name')
        ]
        self.fields['courses'].choices = course_choices


@admin.register(BulkEnrollment)
class BulkEnrollmentAdmin(admin.ModelAdmin):
    
    def has_change_permission(self, request, obj=None):
        return False

    def has_delete_permission(self, request, obj=None):
        return False

    def has_view_permission(self, request, obj=None):
        return False

    def has_module_permission(self, request):
        return True

    def has_add_permission(self, request):
        return True

    def add_view(self, request, form_url='', extra_context=None):
        if request.method == 'POST':
            form = BulkEnrollmentForm(request.POST)
            if form.is_valid():
                emails = [e.strip() for e in form.cleaned_data['emails'].split(',') if e.strip()]
                courses = form.cleaned_data['courses']

                results = {'success': [], 'failed': []}

                for email in emails:
                    for course_id in courses:
                        try:
                            user = User.objects.get(email=email)
                            try:
                                add_enrollment(user.username, course_id, mode=None)
                            except CourseEnrollmentExistsError:
                                pass
                            results['success'].append(f"{email} -> {course_id}")
                        except User.DoesNotExist:
                            results['failed'].append(f"{email} -> {course_id}: User not found")
                        except Exception as e:
                            results['failed'].append(f"{email} -> {course_id}: {e}")

                if results['success']:
                    messages.success(request, f"✅ Success: {len(results['success'])} enrollments")
                if results['failed']:
                    for fail in results['failed']:
                        messages.error(request, f"❌ {fail}")

                return redirect(request.path)
        else:
            form = BulkEnrollmentForm()

        # Use Django's built-in admin template
        from django.contrib.admin.helpers import AdminForm
        
        fieldsets = [(None, {'fields': ['emails', 'courses']})]
        admin_form = AdminForm(
            form,
            fieldsets,
            prepopulated_fields={},
            readonly_fields=[],
            model_admin=self
        )

        context = {
            **self.admin_site.each_context(request),
            'title': 'Bulk Enrollment',
            'adminform': admin_form,
            'form': form,
            'opts': self.model._meta,
            'save_as': False,
            'save_on_top': False,
            'has_add_permission': True,
            'has_change_permission': False,
            'has_delete_permission': False,
            'has_view_permission': False,
            'has_editable_inline_admin_formsets': False,
            'add': True,
            'change': False,
            'is_popup': False,
            'inline_admin_formsets': [],
            'errors': form.errors,
            'show_save': True,
            'show_save_and_continue': False,
            'show_save_and_add_another': False,
            'show_delete': False,
        }
        
        from django.template.response import TemplateResponse
        return TemplateResponse(request, 'admin/change_form.html', context)


##################################################

import io
import csv
from datetime import datetime

from django.contrib import admin
from django.http import HttpResponse
from django import forms

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from openedx.core.djangoapps.content.course_overviews.models import CourseOverview
from common.djangoapps.student.models.course_enrollment import CourseEnrollment
from lms.djangoapps.grades.course_grade_factory import CourseGradeFactory
from opaque_keys.edx.keys import CourseKey


class BulkGradeExport:
    """Dummy class for admin interface"""
    pass


class BulkGradeExportForm(forms.Form):
    """Form for selecting courses"""
    
    courses = forms.MultipleChoiceField(
        widget=forms.SelectMultiple(attrs={'size': 20, 'style': 'width: 600px;'}),
        help_text='Hold Ctrl/Cmd to select multiple courses'
    )
    
    export_format = forms.ChoiceField(
        choices=[
            ('xlsx', 'Excel (.xlsx) - One sheet per course'),
            ('csv', 'CSV (.csv) - All courses in one file'),
        ],
        initial='xlsx',
        widget=forms.RadioSelect
    )
    
    include_inactive = forms.BooleanField(
        required=False,
        initial=False,
        label='Include inactive enrollments'
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        course_choices = [
            (str(course.id), f"{course.display_name} ({course.id})")
            for course in CourseOverview.objects.all().order_by('display_name')
        ]
        self.fields['courses'].choices = course_choices


@admin.register(CourseOverview)
class CourseOverviewGradeAdmin(admin.ModelAdmin):
    """Admin with bulk grade export action"""
    
    list_display = ['display_name', 'id', 'start', 'end', 'enrollment_count']
    list_filter = ['start', 'end']
    search_fields = ['display_name', 'id']
    actions = ['export_grades_xlsx', 'export_grades_csv']
    
    def enrollment_count(self, obj):
        return CourseEnrollment.objects.filter(course_id=obj.id, is_active=True).count()
    enrollment_count.short_description = 'Enrolled Students'
    
    def get_student_grades(self, course_id):
        """Get all students and grades for a course"""
        course_key = CourseKey.from_string(str(course_id))
        enrollments = CourseEnrollment.objects.filter(
            course_id=course_key,
            is_active=True
        ).select_related('user')
        
        data = []
        for enrollment in enrollments:
            user = enrollment.user
            try:
                grade = CourseGradeFactory().read(user, course_key=course_key)
                grade_percent = round(grade.percent * 100, 2)
                letter_grade = grade.letter_grade or 'N/A'
                passed = 'Yes' if grade.passed else 'No'
            except Exception:
                grade_percent = 0
                letter_grade = 'N/A'
                passed = 'N/A'
            
            data.append({
                'username': user.username,
                'email': user.email,
                'full_name': user.get_full_name() or '',
                'enrolled': enrollment.created.strftime('%Y-%m-%d'),
                'grade_percent': grade_percent,
                'letter_grade': letter_grade,
                'passed': passed,
            })
        return data
    
    def sanitize_sheet_name(self, name):
        """Excel sheet names: max 31 chars, no special chars"""
        for char in [':', '\\', '/', '?', '*', '[', ']']:
            name = name.replace(char, '_')
        return name[:31]
    
    @admin.action(description='Export selected courses grades (Excel)')
    def export_grades_xlsx(self, request, queryset):
        """Export grades to Excel with one sheet per course"""
        
        workbook = openpyxl.Workbook()
        workbook.remove(workbook.active)
        
        # Styles
        header_font = Font(bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        headers = ['Username', 'Email', 'Full Name', 'Enrolled Date', 'Grade %', 'Letter', 'Passed']
        col_widths = [18, 35, 25, 15, 10, 10, 10]
        
        for course in queryset:
            sheet_name = self.sanitize_sheet_name(course.display_name)
            ws = workbook.create_sheet(title=sheet_name)
            
            # Course info row
            ws.merge_cells('A1:G1')
            ws['A1'] = f'Course: {course.id}'
            ws['A1'].font = Font(bold=True, size=11)
            
            # Headers
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = thin_border
            
            # Data
            students = self.get_student_grades(course.id)
            for row_idx, student in enumerate(students, 4):
                ws.cell(row=row_idx, column=1, value=student['username']).border = thin_border
                ws.cell(row=row_idx, column=2, value=student['email']).border = thin_border
                ws.cell(row=row_idx, column=3, value=student['full_name']).border = thin_border
                ws.cell(row=row_idx, column=4, value=student['enrolled']).border = thin_border
                ws.cell(row=row_idx, column=5, value=student['grade_percent']).border = thin_border
                ws.cell(row=row_idx, column=6, value=student['letter_grade']).border = thin_border
                ws.cell(row=row_idx, column=7, value=student['passed']).border = thin_border
            
            # Column widths
            for col, width in enumerate(col_widths, 1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
        
        # Response
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        
        filename = f'grades_export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        response = HttpResponse(
            output.read(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        return response
    
    @admin.action(description='Export selected courses grades (CSV)')
    def export_grades_csv(self, request, queryset):
        """Export grades to single CSV file"""
        
        filename = f'grades_export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
        response = HttpResponse(content_type='text/csv')
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        
        writer = csv.writer(response)
        writer.writerow([
            'Course ID', 'Course Name', 'Username', 'Email', 
            'Full Name', 'Enrolled Date', 'Grade %', 'Letter', 'Passed'
        ])
        
        for course in queryset:
            students = self.get_student_grades(course.id)
            for student in students:
                writer.writerow([
                    str(course.id),
                    course.display_name,
                    student['username'],
                    student['email'],
                    student['full_name'],
                    student['enrolled'],
                    student['grade_percent'],
                    student['letter_grade'],
                    student['passed'],
                ])
        
        return response