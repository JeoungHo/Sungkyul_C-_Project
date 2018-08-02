#include "ExcelFormat.h"
#include "ExampleSMPL.h"
#include "Tree_Node.h"
#include <vector>
#include <math.h>
#include <time.h>
#include <string>

#include<stdio.h> 
#include<stdlib.h> 

//#define _DEBUG_
#define _Release_

//#include <cmath>

//const bool KiJang=true;

using namespace ExcelFormat;
using namespace std;

const int No_LAP=1000;

//전역변수 선언
unsigned int no_replication ;					//반복횟수
unsigned int Number_of_Nodes;	//시스템 Tree 구조의 노드 수

vector<Tree_Node> aNode;				//플랜트 트리구조의 모든 노드를 전역변수 벡터로 생성(배열 대신)
vector<Tree_Node> aPrevMaintenance;		//예방정비 )

vector<double> vCumUtil, vAnnualCost, vAnnualRepairTime;

double Total_Cost, CumCost, NPV, CumNPV, interest_rate;

char* LogFileName;
char* WorkBookName;
BasicExcel a;
BasicExcelWorksheet* b; 
BasicExcelCell* c;	
//전역변수 선언 끝

void main(int argc, char* argv[])
{
	//int i;
	clock_t start = clock();
	double duration; 

	double *result;

	double t_end;

	WorkBookName=argv[1];
	if(argc > 3){
		LogFileName=argv[3];
	}
	else{
		LogFileName="LOG.txt";
	}


	//t_end = (double) atoi(argv[1]);

	cout<<"Initializing...\n";

	Number_of_Nodes = Structure_Initialize( &aNode, &aPrevMaintenance, argv[1], &t_end);

	a.New(3 + no_replication * 3);	//출력파일	Sheet 5장 1=로그 2 구성품별 가동율 3=비용 + 가동율/예방정비주기, 4=년도별 가동율, 5=년도별 발생비용, 6=년도별 고장시간

	OutPutInitialize();
	
	result = Simulation_Loop( t_end, no_replication);		//종료시간 365.25일 * 30년
								// t_end					No of Replication		Prev_Interval
	b = a.GetWorksheet(2);
	
	c = b->Cell(1, 0);
	c->SetDouble( result[0] );
	c = b->Cell(1, 1);
	c->SetDouble( result[1] );
		
	duration = (double)(clock() - start) / CLOCKS_PER_SEC;
	c = b->Cell(3, 0);
	c->SetString( "Computation Time(sec.) = " );
	c = b->Cell(3, 1);
	c->SetDouble( duration );

	a.SaveAs(argv[2]);
	a.Close();

	aNode.clear();				//플랜트 트리구조의 모든 노드를 전역변수 벡터로 생성(배열 대신)
	aNode.empty();

	vCumUtil.clear();
	vCumUtil.empty();

	vAnnualCost.clear();
	vAnnualCost.empty();
	vAnnualRepairTime.clear();
	vAnnualRepairTime.empty();
}

double * Simulation_Loop(double t_end, unsigned int No_Rep)
{
	unsigned int i;
	unsigned int repeat;
	int Event, Token;
	unsigned int z;
	//char x;
	//double period;
	char* fname = LogFileName; 	
	//char fname[] = "Log.txt"; 
	FILE *f ; 
	// open test.txt 
	if ( (f = fopen(fname, "w")) == NULL) 
	{ 
		printf("File open error.") ; 
		exit(1) ;	 
	}
	else
	{
		fprintf(f, "No.; Time; Event ID; Facility ID.; Facility Name; Parent ID.; Parent Name; Manpower; Repair Time; ManHour \n");
	}

	int SystemCurrentStatus;
	double SystemCurrentUtil;

	static double result[2];	//re[0] = 시스템 가동율 re[1] = 비용의 현가
	result[0]=result[1]=0.0;
	CumCost = CumNPV = 0;

	vCumUtil.reserve(Number_of_Nodes);		//나중에 통계처리용. 누적 평균 계산
	vCumUtil.assign(Number_of_Nodes, 0.0); 

	vAnnualCost.reserve(Number_of_Nodes);		//나중에 통계처리용. 누적 평균 계산
	vAnnualRepairTime.reserve(Number_of_Nodes);		//나중에 통계처리용. 누적 평균 계산

	for(repeat=0; repeat<No_Rep ; repeat++)
	{
		z=1;

		//시스템 초기화
		smpl();
		//srand((unsigned)repeat);	// for debugging and devlopment
		srand(time(NULL));

		Total_Cost = NPV = 0;

		for(i=0; i<aNode.size() ; i++){
			aNode[i].CurrentUtil=1;
			aNode[i].CumUtil=0;
			aNode[i].PrevTime=0;
			aNode[i].CumCost=0;
			aNode[i].LastCost=0;
			aNode[i].AnnualUtil=0;
		}

		InitializeNodes( 3 );	// 3 = status를 working으로 설정. 2=예방정비, 1=고장, 0=예방정비&고장, -1=대기

		vAnnualCost.assign(Number_of_Nodes, 0.0); 
		vAnnualRepairTime.assign(Number_of_Nodes, 0.0); 

		double PM_Duration = Get_PM_Period(WorkBookName, 1);
		schedule( 3 , 365-PM_Duration , 1 );

			
		//종료 시간 스케쥴
		schedule(100, t_end, -2);

		cout<<"Start Simulation....\n";


		while(time1() <= t_end)
		{
			cause(&Event, &Token);

			//Log 저장
			//Save_Previous_Data(z,Event,Token);
		
			SystemCurrentStatus= aNode[0].status;
			SystemCurrentUtil=aNode[0].CurrentUtil;
			
			switch(Event)
			{
			case 1: /* Failure occurs */
#ifdef _DEBUG_
				if( z%No_LAP == 0)
				{
					printf("%6d,  Time:,%lf *** Event(%d) Facility %02d *** Status %d ->",z, time1(), Event,  Token, aNode[Token].status );
				}
#endif
				status_update_and_schedule_next_event(Event, Token);

				//Update_Current_System_Util();
				fprintf(f, "%d; %lf; %d; %d; %s; %d; %s; %d; %f; %f \n",z, time1(), Event, Token,  aNode[Token].name.c_str(), aNode[Token].ParentNode_ID, aNode[aNode[Token].ParentNode_ID].name.c_str(), aNode[Token].CM_ManPower, aNode[Token].CM_RepairTime, aNode[Token].CM_ManHour);
				
				break;
			case 2: /* 수리 완료 */
#ifdef _DEBUG_
				if( z%No_LAP == 0)
				{
					printf("%6d,  Time:,%lf *** Event(%d) Facility %02d *** Status %d ->",z, time1(), Event,  Token, aNode[Token].status );
				}
#endif

				status_update_and_schedule_next_event(Event, Token);

				break;

			case 3: /* 예방정비 시작 */

				printf("%7d   Time:%lf *** Event(%d) Maintenance %02d *** Begins \n",z, time1(), Event,  Token  );


				//Total_Cost += aPrevMaintenance[Token].FixedCost;
				//NPV += Rounding( aPrevMaintenance[Token].FixedCost * pow( 1 + r, - time1() / 365 ) , 0);
				
				InitializeNodes( -1 );	// 3 = status를 working으로 설정. 2=예방정비, 1=고장, 0=예방정비&고장, -1=대기

				Maintenance(Event, Token);

				PM_Duration = Get_PM_Period(WorkBookName, Token);
				schedule(4, PM_Duration, Token);

				fprintf(f,"%d; %lf; %d; %d; ; ; ; ; %f \n",z, time1(), Event,  Token , PM_Duration);

				break;			
			case 4: /* 예방정비 완료 */

				printf("%7d   Time:%lf *** Event(%d) Maintenance %02d *** Ends \n",z, time1(), Event,  Token  );
				
				InitializeNodes( 3 );	// 3 = status를 working으로 설정. 2=예방정비, 1=고장, 0=예방정비&고장, -1=대기
				//Maintenance(Event, Token);				
				
				PM_Duration = Get_PM_Period(WorkBookName, Token+1);
				schedule( 3 , 365 - PM_Duration , Token+1);

				Update_Current_System_Util();

				Print_Availability_and_Cost(repeat, Token );

				for(i=0 ; i<Number_of_Nodes; i++){
					aNode[i].AnnualUtil = 0.0;
				}

				vAnnualCost.assign(Number_of_Nodes, 0.0); 
				vAnnualRepairTime.assign(Number_of_Nodes, 0.0); 

				fprintf(f,"%d; %lf; %d; %d; ; ; ; ; %f \n",z, time1(), Event,  Token , PM_Duration);

				
				break;

			case 100:	//종료시간
				break;
			}

			Update_Current_System_Util();

#ifdef _DEBUG_
			if( z%No_LAP == 0)
			{
				if(Event <= 2 && Token >=0) printf(" %d \n" , aNode[Token].status );
				else printf(" %d \n" , aNode[0].status);
			}
#endif
			//로그 저장 
			//Save_Changed_Data(z , Event, Token);			
			z++;
		} // end while
		
		printf("PrevPeriod= %d, Replication %d, System Utilization = %f \n" , (int)aPrevMaintenance[0].next_event_time(1), repeat, aNode[0].CumUtil / time1() );
#ifdef _DEBUG_
		for(i=1 ; i < (unsigned)Number_of_Nodes ; i++)
			printf("Node %d Utilization = %f \n", i , aNode[i].CumUtil / time1() );
#endif
		Statistics(repeat+1);	
		fprintf(f, "\nRun %d ends \n\n", repeat + 1);

	} //end for

	SaveStatistics(No_Rep);

	cout << "END of Simulation. SWRO with "<< Number_of_Nodes <<" nodes.\n";
	//cin >> x;

	result[0]= vCumUtil[0] ;
	result[1]= CumNPV ;
	return result;
}

int f_expo(double Lambda1, double null)
{
	return (int)expntl(Lambda1);
}

int f_erlang(double Lambda1, double Lambda2)
{
	return (int)erlang(Lambda1,Lambda2); 
}

int f_uniform(double a, double b)
{
	return (int) uniform (a,b); 
}

int f_weibull(double scale, double shape)
{
	return (int) weibull(scale,shape); 
}

int f_Const(double a, double b)
{
	return (int) a; 
}


void Save_Previous_Data(int row , int Event, int Token){

	//int col; 
		b = a.GetWorksheet(0);
	
		c = b->Cell(row,0);
		c->SetDouble(time1());

		c = b->Cell(row,1);
		c->SetInteger(Event);

		c = b->Cell(row,2);
		c->SetInteger(Token);

	if(Token >= 0){
		c = b->Cell(row,3);
		c->SetInteger(aNode[Token].status);

		c = b->Cell(row,5);
		c->SetDouble(aNode[Token].CurrentUtil);
	}
	else{
		c = b->Cell(row,3);
		c->SetInteger(aNode[0].status);

		c = b->Cell(row,5);
		c->SetDouble(aNode[0].CurrentUtil);
	}

	//a.Save();
}
void Save_Changed_Data(int row , int Event, int Token){

	//BasicExcel a;
	//Workbook.Load("datWorkbook.xls");
	/*BasicExcelWorksheet**/ b = a.GetWorksheet(0);
	//BasicExcelCell* c;	
	//if(Token >= 0){

	//	c = b->Cell(row,4);
	//	c->SetInteger(aNode[Token].status);

	//	c = b->Cell(row,6);
	//	c->SetDouble(aNode[Token].CurrentUtil);
	//	
	//	c = b->Cell(row,7);
	//	c->SetDouble(aNode[Token].LastCost);
	//}
	//else{	//Token < 0 은예방정비를 의미함. 
	//	c = b->Cell(row,4);
	//	c->SetInteger(aNode[0].status);

	//	c = b->Cell(row,6);
	//	c->SetDouble(aNode[0].CurrentUtil);
	//	
	//	c = b->Cell(row,7);
	//	c->SetDouble(aPrevMaintenance[Token].FixedCost);

	//}	
	if(Event == 2){

		c = b->Cell(row,4);
		c->SetInteger(aNode[Token].status);

		c = b->Cell(row,6);
		c->SetDouble(aNode[Token].CurrentUtil);
		
		c = b->Cell(row,7);
		c->SetDouble(aNode[Token].LastCost);
	}
	else if(Event == 4){	//Token < 0 은예방정비를 의미함. 
		c = b->Cell(row,4);
		c->SetInteger(aNode[0].status);

		c = b->Cell(row,6);
		c->SetDouble(aNode[0].CurrentUtil);
		
		c = b->Cell(row,7);
		c->SetDouble(aPrevMaintenance[Token].FixedCost);

	}
	

	//Workbook.Save();
}

int Structure_Initialize(vector<Tree_Node> *node, vector<Tree_Node> *Info, const char* filename, double * t_end)
{
	int  row, node_ID, Parent_ID, structure, required_degree, level, dist1, dist2; 
//	unsigned int i;
	string name;
	double para[2][2];
	Tree_Node tmp_node;
	//int mt[6];
	//unsigned int i;
	double Util_Percent, CostPerDay, FixedCost;
	bool check;
	//2017-10-26
	bool PM_RP_Applicable; 
	unsigned int CM_RP_RFT, PM_RP_Type ;

	unsigned int CM_ManPower; 
	//
	BasicExcel Workbook;

	check = Workbook.Load(filename);
	BasicExcelWorksheet* Sheet = Workbook.GetWorksheet("Structure");
	BasicExcelCell* Cell;	
	
	row = 1; 
	Cell = Sheet->Cell(row,0);				//첫 행, 첫 열 읽음. while 문 조건 때문에... 이후는 while()문 마지막에 있음. 
	node_ID = Cell->GetInteger();
	
	while( row == 1 || node_ID!=0 ) 
	{		
		Cell = Sheet->Cell(row,1);	
		Parent_ID = Cell->GetInteger();
		Cell = Sheet->Cell(row,2);
		name = Cell->GetString();
		Cell = Sheet->Cell(row,3);
		dist1 = Cell->GetInteger();
		Cell = Sheet->Cell(row,4);
		para[0][0] = Cell->GetDouble();
		Cell = Sheet->Cell(row,5);
		para[0][1]= Cell->GetDouble();
		Cell = Sheet->Cell(row,6);
		dist2 = Cell->GetInteger();
		Cell = Sheet->Cell(row,7);
		para[1][0] = Cell->GetDouble();
		Cell = Sheet->Cell(row,8);
		para[1][1]= Cell->GetDouble();
		Cell = Sheet->Cell(row,9);
		structure= Cell->GetInteger();
		Cell = Sheet->Cell(row,10);
		required_degree= Cell->GetInteger();
		Cell = Sheet->Cell(row, 11); 
		Util_Percent = Cell->GetDouble();

		Cell = Sheet->Cell(row,13);
		level= Cell->GetInteger();

		Cell = Sheet->Cell(row, 14); 
		CM_RP_RFT = Cell->GetInteger();
		Cell = Sheet->Cell(row, 15); 
		FixedCost = Cell->GetDouble();
		Cell = Sheet->Cell(row, 16); 
		CostPerDay = Cell->GetDouble();
		Cell = Sheet->Cell(row, 17); 
		CM_ManPower = Cell->GetInteger();
		Cell = Sheet->Cell(row, 18); 
		PM_RP_Type = Cell->GetInteger();
		Cell = Sheet->Cell(row, 19); 
		PM_RP_Applicable = (bool)Cell->GetInteger();

		//노드에 값 대입

		tmp_node.ID = node_ID;
		tmp_node.ParentNode_ID = Parent_ID;
		tmp_node.pParentNode = &(*node)[Parent_ID];
		tmp_node.name = name;
		//tmp_node.set_failure_parameters(dist1*(int)CM_RP_RFT, para[0][0], para[0][1]);
		tmp_node.set_failure_parameters(dist1, para[0][0], para[0][1]);
		tmp_node.set_fixing_parameters(dist2, para[1][0], para[1][1]);
		tmp_node.child_structure=structure;
		tmp_node.no_of_running = required_degree;
		tmp_node.level = level;
		tmp_node.Util_Percent = Util_Percent;
		tmp_node.CostPerDay = CostPerDay;
		tmp_node.FixedCost = FixedCost;
		tmp_node.degree = 0;

		tmp_node.CM_RP_RFT=CM_RP_RFT;
		tmp_node.CM_ManPower=CM_ManPower;

		tmp_node.PM_RP_Type=PM_RP_Type;
		tmp_node.PrevMaintenanceLevel= PM_RP_Type;
		tmp_node.PM_RP_Applicable=PM_RP_Applicable;
		

		node->push_back(tmp_node);

		if(Parent_ID >= 0){
			(*node)[Parent_ID].vChildNodes_ID.push_back(node_ID);
			(*node)[Parent_ID].vChildNodes.push_back( &(*node)[node_ID] );
			(*node)[Parent_ID].degree = (*node)[Parent_ID].vChildNodes_ID.size();
		}

		row ++;
		Cell = Sheet->Cell(row,0);	
		node_ID = Cell->GetInteger();
	}

	//가동 조건, 예방정비 주기 입력
	Sheet = Workbook.GetWorksheet("Info");
	Cell = Sheet->Cell(0, 1);
	*t_end = 365 * Cell->GetDouble();

	Cell = Sheet->Cell(1, 1);
	no_replication = Cell->GetInteger();
	Cell = Sheet->Cell(2, 1);
	interest_rate = Cell->GetDouble();

	//double PrevCycle, PrevDuration, PrevPeriod, PrevCost
	//int PrevNo
	//vector<Tree_Node> *Info

	row = 3; 

	Cell = Sheet->Cell(row, 0);
	tmp_node.ID  = Cell->GetInteger();	//PrevNo
	do
	{
		dist1 = 1;
		Cell = Sheet->Cell(row,3);
		para[0][0] = Cell->GetDouble();
		para[0][1]= 0.0;
		tmp_node.set_failure_parameters(dist1, para[0][0], para[0][1]);

		dist2 = 1;
		Cell = Sheet->Cell(row,2);
		para[1][0] = Cell->GetDouble();
		para[1][1]= 0.0;
		tmp_node.set_fixing_parameters(dist2, para[1][0], para[1][1]);	

		Cell = Sheet->Cell(row, 4);
		tmp_node.FixedCost  = Cell->GetDouble();

		Info->push_back(tmp_node);

		row ++;
		Cell = Sheet->Cell(row, 0);
		tmp_node.ID  = Cell->GetInteger();	//PrevNo
	}while( tmp_node.ID != 0 );

	Workbook.Close();

	return node->size(); //Number of Nodes

}




int AncestorStatusUpdate(int token_no, int event_no)	//상태 결과 리턴
{
	unsigned int i;
	//int ParentID, SiblingID, UpID, ChildID;
	int ParentID, SiblingID, ChildID;
	//bool check;
	//int s, s3, s2, s1, s0, s_1;
	int s, s1;
	double u; 

	if(token_no < 0 ) 
		return -2;

	ParentID = aNode[token_no].ParentNode_ID ;

	// token_no가 최상위 노드가 아니고, 자신이 대기 구조의 일부인 경우,
	if( ParentID >=0 && aNode[ParentID].child_structure == 3)
	{
		if(event_no == 1)	// 고장 발생: 형제 중 대기 -> 가동으로 전환
		{
			for(i=0; i<aNode[ParentID].vChildNodes_ID.size() ; i++){
				SiblingID=aNode[ParentID].vChildNodes_ID[i];

				if(aNode[SiblingID].status == -1 && SiblingID != token_no ){
					SetStatus(SiblingID, 3);
//					Update_Utilization(SiblingID);
					break;
				}
			}
		}
		else if(event_no ==2)	//수리: 부모의 상태에 따라 
		{
			if( abs(aNode[ParentID].CurrentUtil - 1.0) < 0.0000001 ){
				SetStatus(token_no, -1 );
				//Update_Utilization(token_no);
			}

		}//else if(event_no ==2)
	}
	

	switch(aNode[token_no].child_structure)
	{
	case 0:	//leaf node
		cout<<"Error 01 /n";
		return -2;
	case 1:	//serial		//3=가동중 0b11, 2=예방정비 0b10, 1=고장 0b01, 0=고장&예방정비 중 0b00, -1=대기중
		s = 3;
		u=1;
		for(i=0; i< (unsigned)aNode[token_no].degree ; i++){
			ChildID=aNode[token_no].vChildNodes_ID[i];
			if(aNode[ChildID].child_structure == 0 && aNode[ChildID].CM_RP_RFT == 0){ 
				s = min(s, 3);
			}
			else{
				s = min(s, aNode[ChildID].status); 
			}
		}

		break;

	case 2:	//parallel
	case 3:	//Ready
	case 4:	//Multi resource
		s = -2;
		for(i=0; i< (unsigned)aNode[token_no].degree ; i++){

			ChildID=aNode[token_no].vChildNodes_ID[i];
			if(aNode[ChildID].child_structure == 0 && aNode[ChildID].CM_RP_RFT == 0){ 
				s = max(s, 3);
			}
			else{
				s = max(s, aNode[ChildID].status);
			}
		}
		break;
	}
	
	aNode[token_no].set_status( s );
	//u = Update_Utilization(token_no);

	if(ParentID >= 0) // 현재 노드가 최상위 노드가 아니면
		AncestorStatusUpdate(ParentID , event_no );
	
	return s;
}


void status_update_and_schedule_next_event(int event_no, int token_no)
{
	int i, ParentID, SiblingID, status;
	double FixingDays; 

	//구성품[Token] 상태 업데이트
	switch(event_no)
	{
	case 1:	//고장 발생, 고장 발생은 Leaf Node에서만 발생함. 부모 노드는 상태만 변하지 고장 자체가 발생하지는 않음. 
		i=aNode[token_no].status;
		FixingDays = aNode[token_no].set_status(i - 2);		//사건 예약은 이 메소드 내에서 이뤄짐

		ParentID = aNode[token_no].ParentNode_ID;

		//고장 비용 누적
		aNode[token_no].LastCost = Rounding( aNode[token_no].CostPerDay * FixingDays + aNode[token_no].FixedCost , 0);
		
		vAnnualRepairTime[token_no] += FixingDays ;
		vAnnualCost[token_no] += aNode[token_no].LastCost;
		aNode[token_no].CumCost += aNode[token_no].LastCost;
		Total_Cost += Rounding( aNode[token_no].CostPerDay * FixingDays + aNode[token_no].FixedCost , 0);
		NPV += Rounding( aNode[token_no].CostPerDay * FixingDays + aNode[token_no].FixedCost  * pow( 1 + interest_rate, - time1() / 365 ) , 0);
		
		switch(aNode[ParentID].child_structure){
		case 0:	//부모 노드가 Leaf 노드인 경우 -> 그런 경우 없음. 이미 자녀 노드에 고장이 발생했으므로...
			break;
		case 1: //부모 노드가 직렬 구조인 경우, -> 아무 일도 하지 않음 
			break;
		case 2:	//병렬 구조 -> 역시 아무일도 하지 않음. 
			break;
		case 3:	//부모 노드 = 대기 구조 -> 형제 중 대기 중인 노드를 Up
			for(i=0; i< aNode[ParentID].degree; i++){
				SiblingID = aNode[ParentID].vChildNodes_ID[i];
				if( SiblingID != token_no && aNode[SiblingID].status == -1){
			
					SetStatus(SiblingID , 3);	//Ready 구조가 중첩될 경우, 모든 자손 노드들을 up 시켜야 함.
					break;
				}
			}
		case 4:	//복수 자원 -> 아무 일도 하지 않음. 
			break;
		default:
			break;
		}

		break;

	case 2: //고장 수리 완료

		ParentID = aNode[token_no].ParentNode_ID;

		switch(aNode[ParentID].child_structure){
		case 3:	//부모 노드가 Ready 구조인 경우 -> 부모 노드의 상태에 따라 현 컴포넌트의 상태 결정
			switch(aNode[ParentID].status){
			case (-1):	//부모노드 상태 대기 중 -> 아래 Case로 이어짐
			case 3:	// 부모노드 상태 = 대기 중 또는 가동 중 -> 나도 대기 중으로
				if( aNode[ParentID].CurrentUtil >= 0.9999)
					SetStatus(token_no, -1);		//혹시나, 자녀 노드가 있다면 모두 상태를 대기로 전환
				else{
					status = aNode[token_no].status ;
					SetStatus(token_no, status + 2);	
				}
				break;

			case 0:	//부모 노드  고장 + 예방정비 -> 현재 노드 예방정비로 설정
			case 2:	//부모 노드  예방정비 중 -> 현재 노드 예방정비로
				SetStatus(token_no, 2);		//
				break;

			case 1:	//부모 노드 현재 고장 -> 현재 노드 가동으로 설정
				status = aNode[token_no].status ;
				SetStatus(token_no, status + 2);		//
				break;

			default:	//

				break;

			}
			break;

		default:
			i=aNode[token_no].status;
			aNode[token_no].set_status(i + 2);

			break;

		}

		break;

	}

	//Update_Utilization(token_no);

	// 부모 노드 상태 업데이트
	AncestorStatusUpdate(ParentID, event_no);
}

void Statistics(int iteration)
{
	unsigned int row; 

	//b = a.GetWorksheet(1);

	//c = b->Cell(0,col);
	//c->SetInteger(col);

	//for(row=0; row < Number_of_Nodes ; row++){		//이번 replication의 평균 가동율 저장
	//	c = b->Cell(row+1,col);
	//	c->SetDouble(aNode[row].CumUtil / time1() );
	//}
	//c = b->Cell(++row,col);
	//c->SetDouble(Total_Cost );
	//c = b->Cell(++row,col);
	//c->SetDouble( NPV );
		
	unsigned int i;			//가동율 평균

	for(i=0 ; i < Number_of_Nodes ; i++){
		vCumUtil[i] = vCumUtil[i]*(iteration-1)/iteration +  aNode[i].CumUtil / time1() /iteration ;
	}
	CumCost = CumCost *(iteration-1)/iteration + Total_Cost/iteration;
	CumNPV = CumNPV  *(iteration - 1)/iteration  + NPV/iteration ;
}

void SaveStatistics(int replications){

	unsigned int row, col; 
	col = 6;

	b = a.GetWorksheet(1);
	
	c = b->Cell(0, col);
	c->SetString( "Mean" );

	for(row=0; row < Number_of_Nodes ; row++){		//총 replication 평균 가동율 저장
		c = b->Cell(row+1, col);
		c->SetDouble( vCumUtil[row] );
	}
	c = b->Cell(++row, col);
	c->SetDouble( CumCost );
	c = b->Cell(++row, col);
	c->SetDouble( CumNPV );
}

void InitializeNodes(int s)
{
	unsigned int i;

	for(i=0 ; i < aNode.size() ; i++)
	{
		if( aNode[i].ParentNode_ID < 0 )
			SetStatus(i, s);
	}
}


void SetStatus(int ID, int Status)
{
	unsigned int  j;
	int node_ID, structure;
	//double util;

	//if( aNode[ID].Status 

	aNode[ID].set_status( Status );

	structure = aNode[ID].child_structure; 

	switch(structure)
	{
		case 3:	//대기구조	
			if( Status == 3)
			{
				for(j=0; j< (unsigned)aNode[ID].no_of_running; j++){
					node_ID = aNode[ID].vChildNodes_ID[j];

					SetStatus(node_ID, Status); 
				}
				for( ; j<aNode[ID].vChildNodes_ID.size() ; j++){
					node_ID = aNode[ID].vChildNodes_ID[j];
					SetStatus(node_ID, -1);
				}
			}
			else
			{
				for(j=0 ; j<aNode[ID].vChildNodes_ID.size() ; j++){
					node_ID = aNode[ID].vChildNodes_ID[j];
					SetStatus(node_ID, Status);
				}
			}

			break;


		default: //1 직렬, 2 병렬, 4 복수자원
			for(j=0 ; j<aNode[ID].vChildNodes_ID.size() ; j++){
				node_ID = aNode[ID].vChildNodes_ID[j];
				SetStatus(node_ID, Status);
			}

			break;
	}

	/*util = Update_Utilization(ID);
	return util;*/
}

void OutPutInitialize(void)
{
	b = a.GetWorksheet(0);

	c = b->Cell(0, 0);
	c->SetString("Time");
	c = b->Cell(0, 1);
	c->SetString("Event");
	c = b->Cell(0, 2);
	c->SetString("Node");
	c = b->Cell(0, 3);
	c->SetString("Prev Status");
	c = b->Cell(0, 4);
	c->SetString("Cur Status");
	c = b->Cell(0, 5);
	c->SetString("Prev Util");
	c = b->Cell(0, 6);
	c->SetString("Cur Util");
	c = b->Cell(0, 7);
	c->SetString("Cum. Cost");

	//---------------------------------
	b = a.GetWorksheet(1);
	a.RenameWorksheet(1, "Utilization");

	unsigned int row; 
	int p_ID;

	c = b->Cell(0, 0);
	c->SetString("Node ID");
	c = b->Cell(0, 1);
	c->SetString("Node Name");
	c = b->Cell(0, 2);
	c->SetString("PaerentID");
	c = b->Cell(0, 3);
	c->SetString("Parent Name");

	for(row=0; row < Number_of_Nodes ; row++){
		c = b->Cell(row+1, 0);
		c->SetInteger( aNode[row].ID);

		c = b->Cell(row+1, 1);
		const char* string1 =  aNode[row].name.c_str();
		c->SetString(string1);

		p_ID = aNode[row].ParentNode_ID;
		c = b->Cell(row+1, 2);
		c->SetInteger( p_ID);

		if(p_ID >= 0){
			c = b->Cell(row+1, 3);
			const char* string2 =  aNode[p_ID].name.c_str();
			c->SetString(string2);
		}
	}
	c = b->Cell(++row, 0);
	c->SetString("Total_Cost");	
	
	c = b->Cell(++row, 0);
	c->SetString("NPV");

	//--------------------------
	b = a.GetWorksheet(2);
	a.RenameWorksheet(2, "System Utilization and NPV");
	
	c = b->Cell(0, 0);
	c->SetString( "System Utilization" );
	c = b->Cell(0, 1);
	c->SetString( "NPV" );
}

double Update_Current_System_Util( void )
{
	unsigned int i;

	for(i=0 ; i < aNode.size() ; i++)
	{
		if( aNode[i].ParentNode_ID < 0 )
			Update_Utilization( i );
	}

	return aNode[0].CurrentUtil;
}

double Update_Utilization(int tk)
{
	unsigned int i;
	int ChildID;
	double util;
	//누적 가동율 업데이트
	aNode[tk].AnnualUtil +=  (time1() - aNode[tk].PrevTime) *  aNode[tk].CurrentUtil; 
	aNode[tk].CumUtil +=  (time1() - aNode[tk].PrevTime) *  aNode[tk].CurrentUtil; 
	
	aNode[tk].PrevTime = time1();


	//if( aNode[tk].status < 3 ){
	//	aNode[tk].CurrentUtil = 0;
	//	return 0;
	//}

	switch(aNode[tk].child_structure ){
	case 0:	//leaf node
		if(aNode[tk].status == 3) util = 1;
		else util = 0;
		break;
	case 1:	//serial		//3=가동중 0b11, 2=예방정비 0b10, 1=고장 0b01, 0=고장&예방정비 중 0b00, -1=대기중
		util = 1;
		for( i=0 ; i < aNode[tk].vChildNodes_ID.size() ; i++ ){
			ChildID = aNode[tk].vChildNodes_ID[i];
			util = min( util, Update_Utilization(ChildID) );	
		}	
		break;

	case 2:	//parallel
		util = 0;
		for( i=0 ; i < aNode[tk].vChildNodes_ID.size() ; i++ ){
			ChildID = aNode[tk].vChildNodes_ID[i];
			util = max( util, Update_Utilization(ChildID) );	
		}
		break;

	case 3:	//ready 
	case 4:	//복수자원
		util = 0;
		for( i=0 ; i < aNode[tk].vChildNodes_ID.size() ; i++ ){
			ChildID = aNode[tk].vChildNodes_ID[i];
			util +=  Update_Utilization(ChildID) * aNode[ChildID].Util_Percent ;	
		}
		util = min( 1.0, util);
		break;

	}
	util = Rounding( util, 7);
	aNode[tk].CurrentUtil = util;

	if(aNode[tk].child_structure == 0 && aNode[tk].CM_RP_RFT == 0) return 1;
	else return util;

}



double Rounding(double src, int precision) {
         long des;
         double tmp;
         double result;
         tmp = src * pow((double)10, (int)precision);
         if(tmp < 0) {//negative double
            des = (long)(tmp - 0.5);
         }else {
            des = (long)(tmp + 0.5);
         }
         result = (double)((double)des * pow((double)10, (int)-precision));
         return result;
}


void Maintenance(int ev, int tk)
{
	unsigned int row; 
	bool check;
	int NodeID, PM_Level;

	BasicExcel Workbook;

	check = Workbook.Load(WorkBookName);	//WorkBookName은 전역 변수
	BasicExcelWorksheet* Sheet = Workbook.GetWorksheet("PM_Plan");
	BasicExcelCell* Cell;	

	switch( ev )
	{
	case 3:	//예방정비 PM 시작
		row=4 ;
		Cell = Sheet->Cell(row,0);

		while( Cell->Get(NodeID) ) 
		{
			Cell = Sheet->Cell(row,6 + tk);
			PM_Level=Cell->GetInteger();

			if( PM_Level > 0 ) PM_By_Level(NodeID, PM_Level);

			row++;
			Cell = Sheet->Cell(row,0);
		}
		break;
	case 4:	//PM 종료
		//for(row=Number_of_Nodes -1 ; row>0; row--) 
		//{
		//	if(aNode[row].PrevMaintenance[tk] == 1)
		//	{
		//		aNode[i].set_status(aNode[i].status + 1);
		//	}
		//	//else if(aNode[i].status == 3)
		//	//{		
		//	//	aNode[i].set_status( -1);
		//	//	//AncestorStatusUpdate(aNode[i].ParentNode_ID, 3);		
		//	//}
		//	else if(aNode[i].status == -1)
		//	{
		//		aNode[i].set_status( aNode[i].PrevStatus );
		//	}
		//}
		break;
	}
	
	Workbook.Close();
}


void PM_By_Level(int NodeID, unsigned int PM_Level)
{
	unsigned int j;
	int node_ID;
	
	aNode[NodeID].set_status(aNode[NodeID].status - 1);

	for(j=0 ; j<aNode[NodeID].vChildNodes_ID.size() ; j++){
		node_ID = aNode[NodeID].vChildNodes_ID[j];
		PM_By_Level_Below(node_ID, PM_Level);
	}

}


void PM_By_Level_Below(int NodeID, unsigned int PM_Level)
{
	unsigned int j;
	int node_ID;
	
	if(aNode[NodeID].PM_RP_Applicable == true)
	{	if(aNode[NodeID].PM_RP_Type  >= PM_Level)	//해당노드 레벨이 PM 레벨보다 값이 크면, PM 레벨이 낮으므로 PM 실시
		{
			aNode[NodeID].set_status(aNode[NodeID].status - 1);
			for(j=0 ; j<aNode[NodeID].vChildNodes_ID.size() ; j++){
				node_ID = aNode[NodeID].vChildNodes_ID[j];
				PM_By_Level_Below(node_ID, PM_Level);
			}
		}
		else if(aNode[NodeID].status == 3 )
		{
			aNode[NodeID].set_status( -1 );
		}

	}
}

void Print_Availability_and_Cost(int repeat, int year)
{
	unsigned int i, j ;
	//static int column = 4;
	int Parent_ID;

	if(year == 1)
	{
		for(j=1; j<=3 ;j++)
		{
			b = a.GetWorksheet(3 + 3*repeat + j - 1 );

			c = b->Cell(0, 0);
			c->SetString("Node ID");
			for(i=0; i< Number_of_Nodes ; i++){
				c = b->Cell(i+1, 0);
				c->SetInteger(aNode[i].ID);
			}		
		
			c = b->Cell(0, 1);
			c->SetString("Name");
			for(i=0; i< Number_of_Nodes ; i++){
				c = b->Cell(i+1, 1);
				const char* string2 =  aNode[i].name.c_str();
				c->SetString(string2);
			}

			c = b->Cell(0, 2);
			c->SetString("Parent Node ID");
			for(i=0; i< Number_of_Nodes ; i++){
				c = b->Cell(i+1, 2);
				c->SetInteger(aNode[i].ParentNode_ID);
			}		
		
			c = b->Cell(0, 3);
			c->SetString("Parent Name");
			for(i=0; i< Number_of_Nodes ; i++){
				c = b->Cell(i+1, 3);
				Parent_ID=aNode[i].ParentNode_ID;
				if(Parent_ID >=0){
					const char* string1 =  aNode[Parent_ID].name.c_str();
					c->SetString(string1);
				}
			}
		}
	}
	string str; 
	char str2[4];
	sprintf(str2, "%d", repeat + 1);

	str = "Annual Availability " + string(str2);
	const char * sheetname1= str.c_str();
	a.RenameWorksheet(3 + 3*repeat + 1 - 1, sheetname1);

	str = "Annual CM Cost " + string(str2);
	const char * sheetname2= str.c_str();

	a.RenameWorksheet(3 + 3*repeat + 2 - 1, sheetname2);

	str = "Annual CM Hours " + string(str2);
	const char * sheetname3= str.c_str();
	a.RenameWorksheet(3 + 3*repeat + 3 - 1, sheetname3);

	//---- 가용도 출력
	b = a.GetWorksheet(3 + 3*repeat + 1 - 1);

	c = b->Cell(0, year + 3);
	c->SetInteger(year);

	for(i=0; i< Number_of_Nodes ; i++){
		if( aNode[i].AnnualUtil > 0){
			c = b->Cell(i+1, year + 3);
			c->SetDouble( aNode[i].AnnualUtil / 365.0);
		}
	}
	//---- Cost 출력
	b = a.GetWorksheet(3 + 3*repeat + 2 - 1);

	c = b->Cell(0, year + 3);
	c->SetInteger(year);

	for(i=0; i< Number_of_Nodes ; i++){
		if(vAnnualCost[i]){ 
			c = b->Cell(i+1, year + 3);
			c->SetDouble( vAnnualCost[i] );
		}
	}

	//---- Repair Time 출력
	b = a.GetWorksheet(3 + 3*repeat + 3 - 1);

	c = b->Cell(0, year + 3);
	c->SetInteger(year);

	for(i=0; i< Number_of_Nodes ; i++){
		if( vAnnualRepairTime[i] > 0 ){
			c = b->Cell(i+1, year + 3);
			c->SetDouble( vAnnualRepairTime[i] );
		}
	}

}

double Get_PM_Period(char*  FileName, int Column)
{
	bool check;
	BasicExcel Workbook;
	double PM_Duration;

	check = Workbook.Load(FileName);	//WorkBookName은 전역 변수
	BasicExcelWorksheet* Sheet = Workbook.GetWorksheet("PM_Plan");
	BasicExcelCell* Cell;	

	Cell = Sheet->Cell(2,6+Column);
	PM_Duration = Cell->GetDouble();

	return PM_Duration;
}

