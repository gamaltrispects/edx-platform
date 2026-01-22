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