"""
Admin for managing the connection to the Forums backend service.
"""


from django.contrib import admin, messages
from openedx.core.djangoapps.enrollments.api import add_enrollment
from openedx.core.djangoapps.enrollments.data import CourseEnrollmentExistsError
from django.contrib.auth.models import User
from .models import ForumsConfig

admin.site.register(ForumsConfig)




from django import forms

from django.db import models

class BulkEnrollment(models.Model):
    """Dummy model for admin interface"""
    class Meta:
        managed = False  # No database table
        verbose_name = 'Bulk Enrollment'
        verbose_name_plural = 'Bulk Enrollment'


class BulkEnrollmentForm(forms.ModelForm):
    emails = forms.CharField(
        widget=forms.Textarea(attrs={'rows': 4}),
        help_text='Comma-separated: example@web.com, example2@web.com'
    )
    courses = forms.CharField(
        widget=forms.Textarea(attrs={'rows': 4}),
        help_text='Comma-separated: course-1, course-2'
    )

    class Meta:
        model = BulkEnrollment
        fields = []


@admin.register(BulkEnrollment)
class BulkEnrollmentAdmin(admin.ModelAdmin):
    form = BulkEnrollmentForm

    def has_change_permission(self, request, obj=None):
        return False

    def has_delete_permission(self, request, obj=None):
        return False

    def has_view_permission(self, request, obj=None):
        return False

    def add_view(self, request, form_url='', extra_context=None):
        if request.method == 'POST':
            form = BulkEnrollmentForm(request.POST)
            if form.is_valid():
                emails = [e.strip() for e in form.cleaned_data['emails'].split(',') if e.strip()]
                courses = [c.strip() for c in form.cleaned_data['courses'].split(',') if c.strip()]

                results = {'success': [], 'failed': []}

                for email in emails:
                    for course in courses:
                        try:
                            user = User.objects.get(email=email)
                            try:
                                add_enrollment(user.username, course, mode=None)
                            except CourseEnrollmentExistsError:
                                # If the user is already enrolled in the course, do nothing.
                                pass

                            results['success'].append(f"{email} -> {course}")
                        except Exception as e:
                            results['failed'].append(f"{email} -> {course}: {e}")

                # Feedback
                if results['success']:
                    messages.success(request, f"✅ Success: {len(results['success'])} enrollments")
                if results['failed']:
                    for fail in results['failed']:
                        messages.error(request, f"❌ {fail}")

                return redirect(request.path)

        return super().add_view(request, form_url, extra_context)