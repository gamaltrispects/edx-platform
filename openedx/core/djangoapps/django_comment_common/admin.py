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

from django.db import models
from django.contrib import admin
from django.http import HttpResponse
from django import forms
from django.contrib import messages
from django.shortcuts import redirect
from django.template.response import TemplateResponse
from django.contrib.admin.helpers import AdminForm

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from openedx.core.djangoapps.content.course_overviews.models import CourseOverview
from common.djangoapps.student.models.course_enrollment import CourseEnrollment
from lms.djangoapps.grades.course_grade_factory import CourseGradeFactory
from opaque_keys.edx.keys import CourseKey
from lms.djangoapps.certificates.models import GeneratedCertificate

#from common.djangoapps.student.models import CourseEnrollmentAllowed

class GradeExport(models.Model):
    """Dummy model for admin interface"""
    class Meta:
        managed = False
        verbose_name = 'Grade Export'
        verbose_name_plural = 'Grade Export'
        app_label = 'django_comment_common'


class GradeExportForm(forms.Form):
    courses = forms.MultipleChoiceField(
        widget=forms.SelectMultiple(attrs={'size': 20, 'style': 'width: 600px;'}),
        help_text='Hold Ctrl/Cmd to select multiple courses'
    )
    export_format = forms.ChoiceField(
        choices=[('xlsx', 'Excel (.xlsx)'), ('csv', 'CSV (.csv)')],
        initial='xlsx',
        widget=forms.RadioSelect
    )
    include_invited = forms.BooleanField(
        required=False,
        initial=True,
        label='Include invited (not yet enrolled) users'
    )

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        course_choices = [
            (str(course.id), f"{course.display_name} ({course.id})")
            for course in CourseOverview.objects.exclude(catalog_visibility__icontains="none").all().order_by('display_name')
        ]
        self.fields['courses'].choices = course_choices


@admin.register(GradeExport)
class GradeExportAdmin(admin.ModelAdmin):

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
            form = GradeExportForm(request.POST)
            if form.is_valid():
                course_ids = form.cleaned_data['courses']
                export_format = form.cleaned_data['export_format']
                include_invited = form.cleaned_data['include_invited']

                if export_format == 'xlsx':
                    return self.export_xlsx(course_ids, include_invited)
                return self.export_csv(course_ids, include_invited)
        else:
            form = GradeExportForm()

        fieldsets = [(None, {'fields': ['courses', 'export_format', 'include_invited']})]
        admin_form = AdminForm(
            form,
            fieldsets,
            prepopulated_fields={},
            readonly_fields=[],
            model_admin=self
        )

        context = {
            **self.admin_site.each_context(request),
            'title': 'Grade Export',
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

        return TemplateResponse(request, 'admin/change_form.html', context)

    def get_student_grades(self, course_id):
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

                try:
                    if grade.passed and grade.passed_timestamp:
                        passed_date = grade.passed_timestamp.strftime('%Y-%m-%d')
                    else:
                        raise AttributeError("No passed_timestamp")
                except Exception:
                    try:
                        cert = GeneratedCertificate.objects.get(user=user, course_id=course_key)
                        passed_date = cert.created_date.strftime('%Y-%m-%d') if cert.created_date else 'N/A'
                    except Exception:
                        passed_date = 'N/A'

            except Exception:
                grade_percent = 0
                letter_grade = 'N/A'
                passed = 'N/A'
                passed_date = 'N/A'

            data.append({
                'username': user.username,
                'email': user.email,
                'full_name': user.get_full_name() or '',
                'enrolled': enrollment.created.strftime('%Y-%m-%d'),
                'grade_percent': grade_percent,
                'certificate': letter_grade,
                'passed': passed,
                'passed_date': passed_date,
                'status': 'Enrolled',
            })
        return data

    def get_invited_users(self, course_id):
        """Get users who are invited but not yet enrolled."""
        from common.djangoapps.student.models import CourseEnrollmentAllowed
        
        course_key = CourseKey.from_string(str(course_id))
        
        # Get all emails that are already enrolled
        enrolled_emails = set(
            CourseEnrollment.objects.filter(
                course_id=course_key,
                is_active=True
            ).values_list('user__email', flat=True)
        )
        
        # Get invited users who haven't enrolled yet
        invited = CourseEnrollmentAllowed.objects.filter(
            course_id=course_key
        ).exclude(email__in=enrolled_emails)
        
        data = []
        for invite in invited:
            # Check if user exists in the system
            try:
                from django.contrib.auth import get_user_model
                User = get_user_model()
                user = User.objects.get(email=invite.email)
                username = user.username
                full_name = user.get_full_name() or ''
            except User.DoesNotExist:
                username = 'N/A (not registered)'
                full_name = ''
            
            data.append({
                'username': username,
                'email': invite.email,
                'full_name': full_name,
                'enrolled': invite.created.strftime('%Y-%m-%d') if hasattr(invite, 'created') and invite.created else 'N/A',
                'grade_percent': 'N/A',
                'certificate': 'N/A',
                'passed': 'N/A',
                'passed_date': 'N/A',
                'status': 'Invited (Auto-enroll)' if invite.auto_enroll else 'Invited',
            })
        return data

    def sanitize_sheet_name(self, name):
        for char in [':', '\\', '/', '?', '*', '[', ']']:
            name = name.replace(char, '_')
        return name[:31]

    def export_xlsx(self, course_ids, include_invited=True):
        workbook = openpyxl.Workbook()
        workbook.remove(workbook.active)

        header_font = Font(bold=True, color='FFFFFF')
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        invited_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')  # Light yellow for invited
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        headers = ['Username', 'Email', 'Full Name', 'Enrolled/Invited', 'Grade %', 'Certificate', 'Passed', 'Passed Date', 'Status']

        for course_id in course_ids:
            try:
                course = CourseOverview.objects.get(id=course_id)
                sheet_name = self.sanitize_sheet_name(course.display_name)
                course_title = course.display_name
            except CourseOverview.DoesNotExist:
                sheet_name = 'Unknown'
                course_title = course_id

            ws = workbook.create_sheet(title=sheet_name)
            ws.merge_cells('A1:I1')
            ws['A1'] = course_title
            ws['A1'].font = Font(bold=True, size=14)

            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col, value=h)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = thin_border

            # Get enrolled students
            all_data = self.get_student_grades(course_id)
            
            # Add invited users if requested
            if include_invited:
                all_data.extend(self.get_invited_users(course_id))

            for row_idx, s in enumerate(all_data, 4):
                is_invited = s.get('status', '').startswith('Invited')
                for col, key in enumerate(['username', 'email', 'full_name', 'enrolled', 'grade_percent', 'certificate', 'passed', 'passed_date', 'status'], 1):
                    cell = ws.cell(row=row_idx, column=col, value=s[key])
                    cell.border = thin_border
                    if is_invited:
                        cell.fill = invited_fill

            for col, w in enumerate([18, 35, 25, 14, 10, 12, 8, 12, 18], 1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)

        response = HttpResponse(
            output.read(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = f'attachment; filename="grades_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx"'
        return response

    def export_csv(self, course_ids, include_invited=True):
        response = HttpResponse(content_type='text/csv')
        response['Content-Disposition'] = f'attachment; filename="grades_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv"'

        writer = csv.writer(response)
        writer.writerow(['Course ID', 'Course Name', 'Username', 'Email', 'Full Name', 'Enrolled/Invited', 'Grade %', 'Certificate', 'Passed', 'Passed Date', 'Status'])

        for course_id in course_ids:
            try:
                course = CourseOverview.objects.get(id=course_id)
                course_name = course.display_name
            except CourseOverview.DoesNotExist:
                course_name = 'Unknown'

            # Get enrolled students
            all_data = self.get_student_grades(course_id)
            
            # Add invited users if requested
            if include_invited:
                all_data.extend(self.get_invited_users(course_id))

            for s in all_data:
                writer.writerow([
                    course_id,
                    course_name,
                    s['username'],
                    s['email'],
                    s['full_name'],
                    s['enrolled'],
                    s['grade_percent'],
                    s['certificate'],
                    s['passed'],
                    s['passed_date'],
                    s['status'],
                ])

        return response
############




































# class GradeExport(models.Model):
#     """Dummy model for admin interface"""
#     class Meta:
#         managed = False
#         verbose_name = 'Grade Export'
#         verbose_name_plural = 'Grade Export'
#         app_label = 'django_comment_common'


# class GradeExportForm(forms.Form):
#     courses = forms.MultipleChoiceField(
#         widget=forms.SelectMultiple(attrs={'size': 20, 'style': 'width: 600px;'}),
#         help_text='Hold Ctrl/Cmd to select multiple courses'
#     )
#     export_format = forms.ChoiceField(
#         choices=[('xlsx', 'Excel (.xlsx)'), ('csv', 'CSV (.csv)')],
#         initial='xlsx',
#         widget=forms.RadioSelect
#     )

#     def __init__(self, *args, **kwargs):
#         super().__init__(*args, **kwargs)
#         course_choices = [
#             (str(course.id), f"{course.display_name} ({course.id})")
#             for course in CourseOverview.objects.exclude(catalog_visibility__icontains="none").all().order_by('display_name')
#         ]
#         self.fields['courses'].choices = course_choices


# @admin.register(GradeExport)
# class GradeExportAdmin(admin.ModelAdmin):

#     def has_change_permission(self, request, obj=None):
#         return False

#     def has_delete_permission(self, request, obj=None):
#         return False

#     def has_view_permission(self, request, obj=None):
#         return False

#     def has_module_permission(self, request):
#         return True

#     def has_add_permission(self, request):
#         return True

#     def add_view(self, request, form_url='', extra_context=None):
#         if request.method == 'POST':
#             form = GradeExportForm(request.POST)
#             if form.is_valid():
#                 course_ids = form.cleaned_data['courses']
#                 export_format = form.cleaned_data['export_format']

#                 if export_format == 'xlsx':
#                     return self.export_xlsx(course_ids)
#                 return self.export_csv(course_ids)
#         else:
#             form = GradeExportForm()

#         # Use Django's built-in admin template
#         fieldsets = [(None, {'fields': ['courses', 'export_format']})]
#         admin_form = AdminForm(
#             form,
#             fieldsets,
#             prepopulated_fields={},
#             readonly_fields=[],
#             model_admin=self
#         )

#         context = {
#             **self.admin_site.each_context(request),
#             'title': 'Grade Export',
#             'adminform': admin_form,
#             'form': form,
#             'opts': self.model._meta,
#             'save_as': False,
#             'save_on_top': False,
#             'has_add_permission': True,
#             'has_change_permission': False,
#             'has_delete_permission': False,
#             'has_view_permission': False,
#             'has_editable_inline_admin_formsets': False,
#             'add': True,
#             'change': False,
#             'is_popup': False,
#             'inline_admin_formsets': [],
#             'errors': form.errors,
#             'show_save': True,
#             'show_save_and_continue': False,
#             'show_save_and_add_another': False,
#             'show_delete': False,
#         }

#         return TemplateResponse(request, 'admin/change_form.html', context)

#     def get_student_grades(self, course_id):
#         course_key = CourseKey.from_string(str(course_id))
#         enrollments = CourseEnrollment.objects.filter(
#             course_id=course_key,
#             is_active=True
#         ).select_related('user')

#         data = []
#         for enrollment in enrollments:
#             user = enrollment.user
#             try:
#                 grade = CourseGradeFactory().read(user, course_key=course_key)
#                 grade_percent = round(grade.percent * 100, 2)
#                 letter_grade = grade.letter_grade or 'N/A'
#                 passed = 'Yes' if grade.passed else 'No'
#             except Exception:
#                 grade_percent = 0
#                 letter_grade = 'N/A'
#                 passed = 'N/A'

#             data.append({
#                 'username': user.username,
#                 'email': user.email,
#                 'full_name': user.get_full_name() or '',
#                 'enrolled': enrollment.created.strftime('%Y-%m-%d'),
#                 'grade_percent': grade_percent,
#                 'letter_grade': letter_grade,
#                 'passed': passed,
#             })
#         return data

#     def sanitize_sheet_name(self, name):
#         for char in [':', '\\', '/', '?', '*', '[', ']']:
#             name = name.replace(char, '_')
#         return name[:31]

#     def export_xlsx(self, course_ids):
#         workbook = openpyxl.Workbook()
#         workbook.remove(workbook.active)

#         header_font = Font(bold=True, color='FFFFFF')
#         header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
#         thin_border = Border(
#             left=Side(style='thin'),
#             right=Side(style='thin'),
#             top=Side(style='thin'),
#             bottom=Side(style='thin')
#         )
#         headers = ['Username', 'Email', 'Full Name', 'Enrolled', 'Grade %', 'Letter', 'Passed']

#         for course_id in course_ids:
#             try:
#                 course = CourseOverview.objects.get(id=course_id)
#                 sheet_name = self.sanitize_sheet_name(course.display_name)
#                 course_title = course.display_name
#             except CourseOverview.DoesNotExist:
#                 sheet_name = 'Unknown'
#                 course_title = course_id

#             ws = workbook.create_sheet(title=sheet_name)
#             ws.merge_cells('A1:G1')
#             ws['A1'] = course_title
#             ws['A1'].font = Font(bold=True, size=14)

#             for col, h in enumerate(headers, 1):
#                 cell = ws.cell(row=3, column=col, value=h)
#                 cell.font = header_font
#                 cell.fill = header_fill
#                 cell.border = thin_border

#             for row_idx, s in enumerate(self.get_student_grades(course_id), 4):
#                 for col, key in enumerate(['username', 'email', 'full_name', 'enrolled', 'grade_percent', 'letter_grade', 'passed'], 1):
#                     ws.cell(row=row_idx, column=col, value=s[key]).border = thin_border

#             for col, w in enumerate([18, 35, 25, 12, 10, 10, 8], 1):
#                 ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

#         output = io.BytesIO()
#         workbook.save(output)
#         output.seek(0)

#         response = HttpResponse(
#             output.read(),
#             content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
#         )
#         response['Content-Disposition'] = f'attachment; filename="grades_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx"'
#         return response

#     def export_csv(self, course_ids):
#         response = HttpResponse(content_type='text/csv')
#         response['Content-Disposition'] = f'attachment; filename="grades_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv"'

#         writer = csv.writer(response)
#         writer.writerow(['Course ID', 'Course Name', 'Username', 'Email', 'Full Name', 'Enrolled', 'Grade %', 'Letter', 'Passed'])

#         for course_id in course_ids:
#             try:
#                 course = CourseOverview.objects.get(id=course_id)
#                 course_name = course.display_name
#             except CourseOverview.DoesNotExist:
#                 course_name = 'Unknown'

#             for s in self.get_student_grades(course_id):
#                 writer.writerow([
#                     course_id,
#                     course_name,
#                     s['username'],
#                     s['email'],
#                     s['full_name'],
#                     s['enrolled'],
#                     s['grade_percent'],
#                     s['letter_grade'],
#                     s['passed']
#                 ])

#         return response
# ############
