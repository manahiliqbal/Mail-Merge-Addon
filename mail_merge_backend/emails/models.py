from django.db import models

class EmailSchedule(models.Model):
    email_subject = models.CharField(max_length=255)
    email_body = models.TextField()
    schedule_time = models.DateTimeField()
    time_interval = models.IntegerField()
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.email_subject

