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


class BulkGradeExport(CourseOverview):
    """Proxy model to avoid AlreadyRegistered error"""
    
    class Meta:
        proxy = True
        verbose_name = 'Course Grade Export'
        verbose_name_plural = 'Course Grade Exports'


@admin.register(BulkGradeExport)
class BulkGradeExportAdmin(admin.ModelAdmin):
    """Admin with export actions - no custom templates needed"""
    
    list_display = ['display_name', 'id', 'start', 'end', 'get_enrollment_count']
    list_filter = ['start', 'end']
    search_fields = ['display_name', 'id']
    ordering = ['display_name']
    actions = ['export_grades_xlsx', 'export_grades_csv']
    
    # Disable add/delete
    def has_add_permission(self, request):
        return False
    
    def has_delete_permission(self, request, obj=None):
        return False
    
    def get_enrollment_count(self, obj):
        return CourseEnrollment.objects.filter(course_id=obj.id, is_active=True).count()
    get_enrollment_count.short_description = 'Students'
    
    def get_student_grades(self, course_id):
        """Fetch students and grades"""
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
        for char in [':', '\\', '/', '?', '*', '[', ']']:
            name = name.replace(char, '_')
        return name[:31]
    
    @admin.action(description='Export to Excel (one sheet per course)')
    def export_grades_xlsx(self, request, queryset):
        workbook = openpyxl.Workbook()
        workbook.remove(workbook.active)
        
        header_font = Font(bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        headers = ['Username', 'Email', 'Full Name', 'Enrolled', 'Grade %', 'Letter', 'Passed']
        col_widths = [18, 35, 25, 12, 10, 10, 8]
        
        for course in queryset:
            ws = workbook.create_sheet(title=self.sanitize_sheet_name(course.display_name))
            
            # Course header
            ws.merge_cells('A1:G1')
            ws['A1'] = f'Course: {course.display_name} ({course.id})'
            ws['A1'].font = Font(bold=True)
            
            # Column headers
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = thin_border
            
            # Student data
            for row_idx, student in enumerate(self.get_student_grades(course.id), 4):
                for col, key in enumerate(['username', 'email', 'full_name', 'enrolled', 'grade_percent', 'letter_grade', 'passed'], 1):
                    cell = ws.cell(row=row_idx, column=col, value=student[key])
                    cell.border = thin_border
            
            # Column widths
            for col, width in enumerate(col_widths, 1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
        
        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        
        response = HttpResponse(
            output.read(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = f'attachment; filename="grades_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx"'
        return response
    
    @admin.action(description='Export to CSV (all courses)')
    def export_grades_csv(self, request, queryset):
        response = HttpResponse(content_type='text/csv')
        response['Content-Disposition'] = f'attachment; filename="grades_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv"'
        
        writer = csv.writer(response)
        writer.writerow(['Course ID', 'Course Name', 'Username', 'Email', 'Full Name', 'Enrolled', 'Grade %', 'Letter', 'Passed'])
        
        for course in queryset:
            for student in self.get_student_grades(course.id):
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