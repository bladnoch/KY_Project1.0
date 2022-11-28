# #숫자 변환
# test="A"
# test2=1
# for i in range (10):
#     test2=i
#     test3=test+str(test2)
#     print(test3)
# print("")

#영어 변환
tt=['A','B','C','D','E','F','G','H']
tt2=1
for i in range(8):
    tt3=tt[i]+str(tt2)
    print(tt3)

# at=[]
# at2=[]
# for i in range(10): #숫자파트
#     at.append()
#     for j in range(8):

from .models import Article


class ArticleForm(forms.ModelForm):
    title = forms.CharField(

        max_length=100,

        label='제목',

        help_text='제목은 100자이내로 작성하세요.',

        widget=forms.TextInput(

            attrs={

                'class': 'my-input',

                'placeholder': '제목 입력'

            }

        )

    )

    content = forms.CharField(

        label='내용',

        help_text='자유롭게 작성해주세요.',

        widget=forms.Textarea(

            attrs={

                'row': 5,

                'col': 50,

            }

        )

    )

    class Meta:
        model = Article

        fields = ['title', 'content']

