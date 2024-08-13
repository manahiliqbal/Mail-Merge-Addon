from rest_framework import serializers
from .models import EmailSchedule

class EmailScheduleSerializer(serializers.ModelSerializer):
    class Meta:
        model = EmailSchedule
        fields = '__all__'
