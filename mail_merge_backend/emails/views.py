from rest_framework import viewsets
from .models import EmailSchedule
from .serializers import EmailScheduleSerializer

class EmailScheduleViewSet(viewsets.ModelViewSet):
    queryset = EmailSchedule.objects.all()
    serializer_class = EmailScheduleSerializer


